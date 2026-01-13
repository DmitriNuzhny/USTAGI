#!/usr/bin/env python3
"""
Print Monday board column fields (title, id, type) only.

Usage:
  python -m scripts.print_board_fields
  python -m scripts.print_board_fields --board-id 123456789
"""

from __future__ import annotations

import argparse
from dotenv import load_dotenv

from app import _env, fetch_board_schema


def main() -> None:
    load_dotenv()

    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--board-id",
        type=int,
        help="Monday board id (defaults to MONDAY_BOARD_ID env var)",
    )
    args = ap.parse_args()

    token = _env("MONDAY_API_TOKEN")
    board_id = args.board_id or int(_env("MONDAY_BOARD_ID"))

    board = fetch_board_schema(token, board_id)
    cols = board.get("columns") or []

    print("=" * 80)
    print(f"Board: {board.get('name')} ({board.get('id')})")
    print(f"Columns ({len(cols)}):")
    print("=" * 80)

    for c in cols:
        title = c.get("title", "")
        cid = c.get("id", "")
        ctype = c.get("type", "")
        print(f"- {title}  |  id={cid}  |  type={ctype}")

    print("=" * 80)


if __name__ == "__main__":
    main()
