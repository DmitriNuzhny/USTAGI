

#!/usr/bin/env python3
"""
Fetch Monday.com board schema and print ONLY the group
with title 'Ready for Proposal', plus all columns.
"""

import json
import sys
from typing import Any, Dict, List

import requests
from dotenv import load_dotenv
import os

# ======================
# CONFIG
# ======================
BOARD_ID = 8392727910  # <-- PUT YOUR TEST BOARD ID HERE
TARGET_GROUP_TITLE = "Ready for Proposal"

MONDAY_API_URL = "https://api.monday.com/v2"


def monday_graphql(token: str, query: str, variables: Dict[str, Any]) -> Dict[str, Any]:
    headers = {
        "Authorization": token,
        "Content-Type": "application/json",
    }
    resp = requests.post(
        MONDAY_API_URL,
        headers=headers,
        json={"query": query, "variables": variables},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if "errors" in data:
        raise RuntimeError(json.dumps(data["errors"], indent=2))
    return data["data"]


def fetch_board_schema(token: str, board_id: int) -> Dict[str, Any]:
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) {
        id
        name
        groups {
          id
          title
        }
        columns {
          id
          title
          type
          settings_str
        }
      }
    }
    """
    data = monday_graphql(token, query, {"board_id": [board_id]})

    boards = data.get("boards", [])
    if not boards:
        raise RuntimeError(f"No board returned for board_id={board_id}")

    board = boards[0]
    return board


def main():
    load_dotenv()

    token = os.getenv("MONDAY_API_TOKEN", "").strip()
    if not token:
        print("ERROR: MONDAY_API_TOKEN not found in .env file.", file=sys.stderr)
        sys.exit(1)

    board = fetch_board_schema(token, BOARD_ID)

    # ---- Filter groups ----
    matching_groups: List[Dict[str, Any]] = [
        g for g in board.get("groups", [])
        if g.get("title") == TARGET_GROUP_TITLE
    ]

    if not matching_groups:
        print(
            f"ERROR: No group found with title '{TARGET_GROUP_TITLE}'. "
            "Check spelling/case in Monday.",
            file=sys.stderr,
        )
        sys.exit(1)

    result = {
        "board": {
            "id": board["id"],
            "name": board.get("name"),
        },
        "group": matching_groups[0],  # exactly one expected
        "columns": board.get("columns", []),
    }

    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
