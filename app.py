import os
import json
import time
import logging

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

# -----------------------
# Setup
# -----------------------
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger("monday-export-eob")

app = FastAPI()

# Simple in-memory idempotency to ignore repeated deliveries of the same action
# (good enough for local dev; swap to Redis if you deploy multiple workers)
_seen_actions: dict[str, float] = {}
_SEEN_TTL_SECONDS = 60 * 60  # 1 hour


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


def _find_item_id(payload: dict) -> int | None:
    # Most common monday payload paths for app actions
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

    # Recursive fallback
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


def _read_sample_file(path: str) -> tuple[bytes, str]:
    if not os.path.isfile(path):
        raise RuntimeError(f"Sample file not found: {path}")
    with open(path, "rb") as f:
        return f.read(), os.path.basename(path)


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

    # monday's /v2/file expects item_id/column_id in the query; only $file is a variable
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
        headers={"Authorization": api_token},  # don't set Content-Type manually
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


def _get_action_uuid(body: dict) -> str | None:
    try:
        return str(body.get("runtimeMetadata", {}).get("actionUuid") or "").strip() or None
    except Exception:
        return None


def _seen_action(action_uuid: str) -> bool:
    now = time.time()
    # prune old
    for k, ts in list(_seen_actions.items()):
        if now - ts > _SEEN_TTL_SECONDS:
            _seen_actions.pop(k, None)

    if action_uuid in _seen_actions:
        return True

    _seen_actions[action_uuid] = now
    return False


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
        # Already processed this exact action delivery
        return JSONResponse(status_code=200, content={"ok": True, "deduped": True, "actionUuid": action_uuid})

    item_id = _find_item_id(body)
    if not item_id:
        # Permanent bad request: stop retries
        return JSONResponse(status_code=400, content={"error": "missing_item_id"})

    try:
        api_token = _env("MONDAY_API_TOKEN")
        file_column_id = _env("MONDAY_FILE_COLUMN_ID")
        sample_path = _env("SAMPLE_FILE_PATH")

        file_bytes, filename = _read_sample_file(sample_path)

        upload_result = _monday_upload_file_to_column(
            api_token=api_token,
            item_id=item_id,
            column_id=file_column_id,
            file_bytes=file_bytes,
            filename=filename,
        )

        log.info(f"Uploaded sample '{filename}' to item {item_id} column {file_column_id}")
        return JSONResponse(
            status_code=200,
            content={"ok": True, "uploaded": True, "itemId": item_id, "filename": filename, "actionUuid": action_uuid},
        )

    except RuntimeError as e:
        # Treat as permanent (config/sample/monday response): stop monday retry storms
        log.exception(f"Export/upload failed: {e}")
        return JSONResponse(
            status_code=400,
            content={"ok": False, "uploaded": False, "itemId": item_id, "actionUuid": action_uuid, "error": str(e)},
        )
    except Exception as e:
        # Unexpected: you may want retries while debugging true transient issues
        log.exception(f"Unexpected failure: {e}")
        return JSONResponse(
            status_code=500,
            content={"ok": False, "uploaded": False, "itemId": item_id, "actionUuid": action_uuid, "error": str(e)},
        )


if __name__ == "__main__":
    import uvicorn

    # run: python app.py
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
