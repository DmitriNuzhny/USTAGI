#!/usr/bin/env python3
"""
Local Monday.com development CLI for USTAGI EOB tool.

Usage:
  python -m scripts.monday_local --list-items
  python -m scripts.monday_local --item-id 123 --out outputs/Name__123__estimator.xlsx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from pathlib import Path

from dotenv import load_dotenv

# Import from app.py (production webhook code)
from app import (
    _env,
    _monday_graphql,
    fetch_board_schema,
    build_colid_to_title,
    fetch_item_column_values,
    monday_item_to_inputs,
    decide_mode,
    normalize_inputs_for_mode,
    generate_excel,
)
from dropbox_uploader import upload_eob_workbook


def fetch_item_name(token: str, item_id: int) -> str:
    query = """
    query ($item_id: [ID!]!) {
      items(ids: $item_id) {
        name
      }
    }
    """
    data = _monday_graphql(token, query, {"item_id": [int(item_id)]})
    items = data.get("items") or []
    if not items:
        raise RuntimeError(f"No item returned for item_id={item_id}")
    return items[0].get("name") or ""


def list_board_items(token: str, board_id: int) -> list[dict]:
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) {
        items_page(limit: 100) {
          items { id name }
        }
      }
    }
    """
    data = _monday_graphql(token, query, {"board_id": [int(board_id)]})
    boards = data.get("boards") or []
    if not boards:
        raise RuntimeError(f"No board returned for board_id={board_id}")
    items = boards[0].get("items_page", {}).get("items", []) or []
    return [{"id": int(it["id"]), "name": it.get("name") or ""} for it in items]


def main() -> None:
    load_dotenv()

    ap = argparse.ArgumentParser(prog="monday_local", description="Local Monday.com development CLI")
    ap.add_argument("--list-items", action="store_true", help="List items on the test board")
    ap.add_argument("--item-id", type=int, help="Item ID to process locally")
    ap.add_argument("--out", help="Output path (must be under outputs/). If omitted, auto-generates.")
    ap.add_argument("--upload-dropbox", action="store_true", help="After generating, upload to Dropbox using current field inputs")

    args = ap.parse_args()

    if not args.list_items and not args.item_id:
        ap.error("Must specify --list-items or --item-id")

    token = _env("MONDAY_API_TOKEN")
    board_id = int(_env("MONDAY_BOARD_ID"))

    if args.list_items:
        items = list_board_items(token, board_id)
        for it in items:
            print(f"{it['name']} : {it['id']}")
        return

    # --item-id mode
    item_id = args.item_id
    item_name = fetch_item_name(token, item_id)
    if not item_name:
        raise RuntimeError(f"Item {item_id} has no name")

    safe_name = item_name.replace(" ", "_").replace("/", "_")

    # Fetch inputs for generation
    board_schema = fetch_board_schema(token, board_id)
    colid_to_title = build_colid_to_title(board_schema)
    col_vals = fetch_item_column_values(token, item_id)
    field_inputs = monday_item_to_inputs(col_vals, colid_to_title, item_name=item_name)
    mode = decide_mode(field_inputs)
    normalize_inputs_for_mode(field_inputs, mode)

    if mode == "residential" and "Tier" not in field_inputs:
        field_inputs["Tier"] = "SFR$$"

    # Generate filename from property address
    prop_addr_raw = field_inputs.get("Property Address") or "Unknown Address"
    # Sanitize for filename (remove invalid chars)
    prop_addr_safe = re.sub(r'[<>:"/\\|?*]', '', str(prop_addr_raw)).strip()
    if not prop_addr_safe:
        prop_addr_safe = "Unknown Address"

    if args.out:
        out_path = Path(args.out)
        if not out_path.is_relative_to(Path("outputs")):
            raise RuntimeError("Output path must be under outputs/")
    else:
        # Auto-generate with new naming convention
        out_path = Path("outputs") / f"EOB {prop_addr_safe}.xlsx"

    Path("outputs").mkdir(exist_ok=True)
    generate_excel(mode, field_inputs, out_path)
    print(f"Generated {out_path}")

    if args.upload_dropbox:
        file_bytes = out_path.read_bytes()
        filename = out_path.name

        client_name = field_inputs.get("Name") or field_inputs.get("Client Name") or field_inputs.get("Client") or "Unknown Client"
        prop_addr = field_inputs.get("Property Address") or "Unknown Address"

        def _extract_year(val):
            if val is None:
                return None
            if isinstance(val, (int, float)):
                return str(int(val))
            s = str(val).strip()
            if len(s) >= 4 and s[:4].isdigit():
                return s[:4]
            return None

        year = (
            _extract_year(field_inputs.get("Date Placed in Service"))
            or _extract_year(field_inputs.get("In-Service Date"))
            or _extract_year(field_inputs.get("Tax Year of CSS"))
            or _extract_year(field_inputs.get("Study Tax Year"))
            or _extract_year(field_inputs.get("Tax Year"))
            or str(__import__("datetime").date.today().year)
        )

        print(f"Dropbox attempt: client={client_name} year={year} address={prop_addr}")
        try:
            upload_eob_workbook(
                file_bytes=file_bytes,
                filename=filename,
                client_name=str(client_name),
                year=str(year),
                property_address=str(prop_addr),
                logger=type("TmpLog", (), {"info": print, "warning": print})(),
            )
        except Exception as e:
            print(f"Dropbox upload skipped/failed: {repr(e)}")


if __name__ == "__main__":
    main()