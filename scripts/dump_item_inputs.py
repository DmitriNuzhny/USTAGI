#!/usr/bin/env python3
"""
Dump Monday item -> normalized estimator inputs, and report missing fields.

Usage:
  python -m scripts.dump_item_inputs --list-items
  python -m scripts.dump_item_inputs --item-id 123
  python -m scripts.dump_item_inputs --item-id 123 --json outputs/item_123_inputs.json

Env:
  MONDAY_API_TOKEN
  MONDAY_BOARD_ID
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from dotenv import load_dotenv

# Reuse production code paths from app.py (same as monday_local.py)
from app import (
    _env,
    _monday_graphql,
    fetch_board_schema,
    build_colid_to_title,
    fetch_item_column_values,
    monday_item_to_inputs,
    decide_mode,
    normalize_inputs_for_mode,
)

# ------------------------
# Helpers
# ------------------------

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


def required_fields_for_mode(mode: str) -> list[str]:
    """
    Minimal fields to generate a *non-empty* schedule in the Python engine.
    (You can expand this list later.)
    """
    if mode == "residential":
        return [
            "Basis",
            "Date Placed in Service",
            "Study Tax Year",
            "Tier",
        ]
    # commercial
    return [
        "Basis",
        "In-Service Date",
        "Study Tax Year",
        "Property Type",
    ]


def find_missing(inputs: dict, required: list[str]) -> list[str]:
    missing = []
    for k in required:
        v = inputs.get(k)
        if v is None:
            missing.append(k)
            continue
        if isinstance(v, str) and not v.strip():
            missing.append(k)
            continue
    return missing


def main() -> None:
    load_dotenv()

    ap = argparse.ArgumentParser(prog="dump_item_inputs", description="Dump Monday item -> estimator inputs")
    ap.add_argument("--list-items", action="store_true", help="List items on the test board")
    ap.add_argument("--item-id", type=int, help="Item ID to inspect")
    ap.add_argument("--json", help="Optional path to write the inputs dict as JSON")

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

    item_id = int(args.item_id)
    item_name = fetch_item_name(token, item_id)

    # Build schema mapping (col id -> title), then fetch values, then normalize
    board_schema = fetch_board_schema(token, board_id)
    colid_to_title = build_colid_to_title(board_schema)
    col_vals = fetch_item_column_values(token, item_id)
    field_inputs = monday_item_to_inputs(col_vals, colid_to_title)

    mode = decide_mode(field_inputs)
    normalize_inputs_for_mode(field_inputs, mode)

    if mode == "residential" and "Tier" not in field_inputs:
        field_inputs["Tier"] = "SFR$$"
        print("[INFO] Tier missing from Monday fields, using default: 'SFR$$'")

    required = required_fields_for_mode(mode)
    missing = find_missing(field_inputs, required)

    # Pretty print
    print("=" * 80)
    print(f"Item: {item_name} ({item_id})")
    print(f"Mode: {mode}")
    print("-" * 80)
    print("Normalized Inputs (what the estimator actually uses):")
    print(json.dumps(field_inputs, indent=2, ensure_ascii=False, default=str))
    print("-" * 80)
    print("Required (minimal) fields for non-empty schedule:")
    for k in required:
        print(f"  - {k}: {field_inputs.get(k)!r}")
    print("-" * 80)
    if missing:
        print("❌ Missing required fields:")
        for k in missing:
            print(f"  - {k}")
        print("\nThis is why you get an empty schedule / zeros (and previously #REF! leaks).")
    else:
        print("✅ Required fields present. If output is still blank/zero, the issue is in calc logic or writing.")
    print("=" * 80)

    if args.json:
        out = Path(args.json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(field_inputs, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        print(f"Wrote inputs JSON to: {out}")


if __name__ == "__main__":
    main()
