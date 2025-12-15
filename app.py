import os
import json
import time
import logging
import tempfile
from pathlib import Path

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

# EOB engine imports
from eob_tool.io import load_inputs
from eob_tool.residential import compute_residential
from eob_tool.commercial import compute_commercial
from eob_tool.main import load_commercial_guidelines_df
from eob_tool.excel_writer import write_residential_workbook, write_commercial_workbook

# -----------------------
# Setup
# -----------------------
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger("monday-export-eob")

app = FastAPI()

# Simple in-memory idempotency to ignore repeated deliveries of the same action
_seen_actions: dict[str, float] = {}
_SEEN_TTL_SECONDS = 60 * 60  # 1 hour

MONDAY_API_URL = "https://api.monday.com/v2"


def _env(name: str) -> str:
    val = os.getenv(name, "").strip()
    if not val:
        raise RuntimeError(f"{name} missing/invalid")
    return val


def _safe_json(obj) -> str:
    try:
        return json.dumps(obj, indent=2)
    except Exception:
        return str(obj)


def _get_action_uuid(body: dict) -> str | None:
    try:
        return str(body.get("runtimeMetadata", {}).get("actionUuid") or "").strip() or None
    except Exception:
        return None


def _seen_action(action_uuid: str) -> bool:
    now = time.time()
    for k, ts in list(_seen_actions.items()):
        if now - ts > _SEEN_TTL_SECONDS:
            _seen_actions.pop(k, None)

    if action_uuid in _seen_actions:
        return True

    _seen_actions[action_uuid] = now
    return False


def _find_item_id(payload: dict) -> int | None:
    candidates = [
        ("payload", "inputFields", "itemId"),
        ("payload", "inboundFieldValues", "itemId"),
        ("payload", "itemId"),
        ("event", "itemId"),
        ("data", "itemId"),
    ]

    for path in candidates:
        cur = payload
        ok = True
        for k in path:
            if isinstance(cur, dict) and k in cur:
                cur = cur[k]
            else:
                ok = False
                break
        if ok:
            try:
                return int(cur)
            except Exception:
                pass

    def walk(x):
        if isinstance(x, dict):
            for k, v in x.items():
                if str(k).lower() in ("itemid", "item_id"):
                    yield v
                yield from walk(v)
        elif isinstance(x, list):
            for v in x:
                yield from walk(v)

    for v in walk(payload):
        try:
            return int(v)
        except Exception:
            continue

    return None


# -----------------------
# monday GraphQL helpers
# -----------------------
def _monday_graphql(token: str, query: str, variables: dict) -> dict:
    resp = requests.post(
        MONDAY_API_URL,
        headers={"Authorization": token, "Content-Type": "application/json"},
        json={"query": query, "variables": variables},
        timeout=60,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("errors"):
        raise RuntimeError(f"monday graphql errors: {data['errors']}")
    return data["data"]


def fetch_item_column_values(token: str, item_id: int) -> list[dict]:
    query = """
    query ($item_id: [ID!]!) {
      items(ids: $item_id) {
        id
        name
        column_values {
          id
          text
          value
          type
        }
      }
    }
    """
    data = _monday_graphql(token, query, {"item_id": [int(item_id)]})
    items = data.get("items") or []
    if not items:
        raise RuntimeError(f"No item returned for item_id={item_id}")
    return items[0].get("column_values") or []


def fetch_board_schema(token: str, board_id: int) -> dict:
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) {
        id
        name
        columns {
          id
          title
          type
          settings_str
        }
      }
    }
    """
    data = _monday_graphql(token, query, {"board_id": [int(board_id)]})
    boards = data.get("boards") or []
    if not boards:
        raise RuntimeError(f"No board returned for board_id={board_id}")
    return boards[0]


def build_colid_to_title(board_schema: dict) -> dict[str, str]:
    m: dict[str, str] = {}
    for c in (board_schema.get("columns") or []):
        m[c["id"]] = c.get("title") or c["id"]
    return m


# -----------------------
# Monday -> EOB input mapping (your board, v1)
# -----------------------
def monday_item_to_inputs(column_values: list[dict], colid_to_title: dict[str, str]) -> dict:
    """
    FIELD-keyed inputs for eob_tool.io:
      - Property Type  (from board column "Property Type")
      - Basis          (from board column "Building Basis")
      - In-Service Date (from board column "In Service Date")
      - Study Tax Year  (from board column "Tax Year of CSS")
      - Bed Cnt (Residential only): "Closed Room Qty"
    Missing/blank values are omitted.
    """
    out: dict = {}

    by_title: dict[str, dict] = {}
    for cv in column_values:
        col_id = cv.get("id")
        title = colid_to_title.get(col_id, "")
        by_title[title.strip().lower()] = cv

    def take(title: str) -> str | None:
        cv = by_title.get(title.lower())
        if not cv:
            return None
        txt = (cv.get("text") or "").strip()
        return txt if txt else None

    # Core fields
    prop_type = take("Property Type")
    if prop_type:
        out["Property Type"] = prop_type

    prop_use = take("Property Use")
    if prop_use:
        out["Property Use"] = prop_use

    basis = take("Building Basis")
    if basis:
        out["Basis"] = basis

    # Lookback inputs (only runs if BOTH exist; calculators enforce)
    isd = take("In Service Date")
    tax_year = take("Tax Year of CSS")
    if isd:
        out["In-Service Date"] = isd
    if tax_year:
        out["Study Tax Year"] = tax_year

    # Residential confirmed mapping
    closed_rooms = take("Closed Room Qty")
    if closed_rooms:
        out["Bed Cnt"] = closed_rooms

    # Nice-to-have for template header if you want it later:
    addr = take("Property Address")
    if addr:
        out["Property Address"] = addr

    return out


def decide_mode(inputs: dict) -> str:
    use = (inputs.get("Property Use") or "").strip().lower()
    if use in ("residential", "res"):
        return "residential"
    if use in ("commercial", "com"):
        return "commercial"
    # fallback if blank/unexpected
    return "commercial"


def generate_excel(mode: str, field_inputs: dict, out_path: Path) -> None:
    """
    Uses your existing io + calculators + excel_writer.
    Converts FIELD-keyed -> cell-keyed by writing a temp json and calling load_inputs().
    """
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as tf:
        json.dump(field_inputs, tf)
        temp_json = tf.name

    B = load_inputs(mode, temp_json)

    if mode == "residential":
        res_payload = compute_residential(B)
        write_residential_workbook(res_payload, out_path)
    else:
        guidelines_df = load_commercial_guidelines_df()
        com_payload = compute_commercial(B, guidelines_df)
        write_commercial_workbook(com_payload, out_path)


def _monday_upload_file_to_column(
    api_token: str,
    item_id: int,
    column_id: str,
    file_bytes: bytes,
    filename: str,
) -> dict:
    """
    monday upload format for https://api.monday.com/v2/file
    multipart fields: query, map, image
    """
    url = "https://api.monday.com/v2/file"

    query = (
        'mutation ($file: File!) { '
        f'add_file_to_column(item_id: {int(item_id)}, column_id: "{column_id}", file: $file) '
        "{ id } }"
    )

    files = {
        "query": (None, query),
        "map": (None, json.dumps({"image": "variables.file"})),
        "image": (filename, file_bytes),
    }

    resp = requests.post(
        url,
        headers={"Authorization": api_token},
        files=files,
        timeout=60,
    )

    try:
        payload = resp.json()
    except Exception:
        raise RuntimeError(f"monday upload: non-JSON response {resp.status_code}: {resp.text[:500]}")

    if resp.status_code >= 400:
        raise RuntimeError(f"monday upload HTTP {resp.status_code}: {payload}")

    if payload.get("errors"):
        raise RuntimeError(f"monday upload returned errors: {payload}")

    return payload


# -----------------------
# Routes
# -----------------------
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/monday/webhook/export-eob")
async def export_eob_webhook(request: Request):
    body = await request.json()
    log.info(f"EXPORT-EOB WEBHOOK: {_safe_json(body)}")

    action_uuid = _get_action_uuid(body)
    if action_uuid and _seen_action(action_uuid):
        return JSONResponse(status_code=200, content={"ok": True, "deduped": True, "actionUuid": action_uuid})

    item_id = _find_item_id(body)
    if not item_id:
        return JSONResponse(status_code=400, content={"error": "missing_item_id"})

    try:
        api_token = _env("MONDAY_API_TOKEN")
        file_column_id = _env("MONDAY_FILE_COLUMN_ID")
        board_id = int(_env("MONDAY_BOARD_ID"))

        board = fetch_board_schema(api_token, board_id)
        colid_to_title = build_colid_to_title(board)

        col_vals = fetch_item_column_values(api_token, item_id)
        field_inputs = monday_item_to_inputs(col_vals, colid_to_title)
        mode = decide_mode(field_inputs)

        log.info(f"Field inputs: {field_inputs}")
        log.info(f"Mode: {mode}")

        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / f"eob_{mode}_item_{item_id}.xlsx"
            generate_excel(mode, field_inputs, out_path)
            file_bytes = out_path.read_bytes()
            filename = out_path.name

        _monday_upload_file_to_column(
            api_token=api_token,
            item_id=item_id,
            column_id=file_column_id,
            file_bytes=file_bytes,
            filename=filename,
        )

        log.info(f"Uploaded '{filename}' to item {item_id} column {file_column_id}")
        return JSONResponse(
            status_code=200,
            content={"ok": True, "uploaded": True, "itemId": item_id, "filename": filename, "mode": mode, "actionUuid": action_uuid},
        )

    except RuntimeError as e:
        log.exception(f"Export/upload failed: {e}")
        return JSONResponse(
            status_code=400,
            content={"ok": False, "uploaded": False, "itemId": item_id, "actionUuid": action_uuid, "error": str(e)},
        )
    except Exception as e:
        log.exception(f"Unexpected failure: {e}")
        return JSONResponse(
            status_code=500,
            content={"ok": False, "uploaded": False, "itemId": item_id, "actionUuid": action_uuid, "error": str(e)},
        )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
