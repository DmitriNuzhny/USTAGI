#!/usr/bin/env python3
"""
dropbox_list_folders.py

- Lists folders/files you can access in Dropbox using API v2.
- Optional: create a test root folder.
- Works with a standard user-scoped token.

Usage:
  export DROPBOX_ACCESS_TOKEN="sl...."
  python dropbox_list_folders.py

Optional env vars:
  DROPBOX_LIST_PATH="/"         # Dropbox path to list ("/" or "/CostSeg Team Folder" etc)
  DROPBOX_MAX_DEPTH="2"         # how deep to recurse folders
  DROPBOX_CREATE_FOLDER="/CostSeg Test"   # if set, create this folder first
"""

import os
from dotenv import load_dotenv

load_dotenv()

import json
import requests

API = "https://api.dropboxapi.com/2"


def get_root_namespace_id(token: str) -> str | None:
    r = requests.post(
        f"{API}/users/get_current_account",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    return data.get("root_info", {}).get("root_namespace_id")


def dbx_post(token: str, endpoint: str, payload: dict, root_ns: str | None):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    if root_ns:
        headers["Dropbox-API-Path-Root"] = json.dumps({
            ".tag": "root",
            "root": root_ns,
        })

    return requests.post(
        f"{API}{endpoint}",
        headers=headers,
        data=json.dumps(payload),
        timeout=60,
    )


def create_folder(token: str, path: str, root_ns: str | None):
    # idempotent-ish: if exists, Dropbox returns 409 conflict(folder) – we treat as OK
    r = dbx_post(token, "/files/create_folder_v2", {"path": path, "autorename": False}, root_ns)
    if r.status_code == 409:
        try:
            err = r.json()
        except Exception:
            return
        conflict = (
            err.get("error", {})
            .get("path", {})
            .get("conflict", {})
            .get(".tag")
        )
        if conflict == "folder":
            print(f"[OK] Folder already exists: {path}")
            return
        raise RuntimeError(f"Create folder conflict: {err}")
    if r.status_code >= 400:
        raise RuntimeError(f"Create folder failed {r.status_code}: {r.text}")
    print(f"[OK] Created folder: {path}")


def list_folder(token: str, path: str, root_ns: str | None):
    # Dropbox uses "" to mean root in list_folder
    dbx_path = "" if path in ("/", "") else path
    r = dbx_post(
        token,
        "/files/list_folder",
        {
            "path": dbx_path,
            "recursive": False,
            "include_deleted": False,
            "include_media_info": False,
            "include_non_downloadable_files": True,
        },
        root_ns,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"list_folder failed {r.status_code}: {r.text}")
    return r.json()


def list_folder_continue(token: str, cursor: str):
    r = dbx_post(token, "/files/list_folder/continue", {"cursor": cursor})
    if r.status_code >= 400:
        raise RuntimeError(f"list_folder/continue failed {r.status_code}: {r.text}")
    return r.json()


def walk_folders(token: str, start_path: str, max_depth: int, root_ns: str | None):
    """
    Depth-limited folder walk: prints folders and files.
    """
    def _walk(path: str, depth: int):
        data = list_folder(token, path, root_ns)
        entries = data.get("entries", [])

        # paginate if needed
        while data.get("has_more"):
            data = list_folder_continue(token, data["cursor"], root_ns)
            entries.extend(data.get("entries", []))

        # print items
        indent = "  " * depth
        folders = []
        for e in entries:
            tag = e.get(".tag")
            name = e.get("name")
            disp = e.get("path_display") or e.get("path_lower")
            if tag == "folder":
                print(f"{indent}[DIR]  {name}   -> {disp}")
                folders.append(disp)
            else:
                print(f"{indent}[FILE] {name}  ({tag})")

        # recurse
        if depth >= max_depth:
            return
        for f in folders:
            _walk(f, depth + 1)

    _walk(start_path, 0)


def main():
    token = os.getenv("DROPBOX_ACCESS_TOKEN", "").strip()
    if not token:
        raise SystemExit("Missing env var: DROPBOX_ACCESS_TOKEN")

    root_ns = get_root_namespace_id(token)

    list_path = os.getenv("DROPBOX_LIST_PATH", "/").strip() or "/"
    max_depth = int(os.getenv("DROPBOX_MAX_DEPTH", "1"))
    create_path = os.getenv("DROPBOX_CREATE_FOLDER", "").strip()

    # Optional: create a folder first (like /CostSeg Test)
    if create_path:
        create_folder(token, create_path, root_ns)

    print(f"\nListing from: {list_path}  (max depth={max_depth})\n")
    walk_folders(token, list_path, max_depth, root_ns)

    print("\nSuggested root examples you can use:")
    print('  DROPBOX_BASE_ROOT="/CostSeg Test"')
    print('  DROPBOX_ALLOWED_ROOT="/CostSeg Test"\n')


if __name__ == "__main__":
    main()
