#!/usr/bin/env python3
"""
dropbox_uploader.py

Dropbox API v2 helper:
- Create nested folders (idempotent)
- Upload file bytes to a target folder

Uses:
- https://api.dropboxapi.com/2/files/create_folder_v2
- https://content.dropboxapi.com/2/files/upload
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from typing import Any, Optional

import requests

RPC = "https://api.dropboxapi.com/2"
CONTENT = "https://content.dropboxapi.com/2"

ILLEGAL = re.compile(r"[\/\0]")


class DropboxError(RuntimeError):
    pass


def _norm_path(p: str) -> str:
    p = (p or "").strip()
    if not p:
        return "/"
    if not p.startswith("/"):
        p = "/" + p
    p = re.sub(r"/{2,}", "/", p)
    if len(p) > 1 and p.endswith("/"):
        p = p[:-1]
    return p


def sanitize_component(s: Any) -> str:
    s = str(s or "").strip()
    s = ILLEGAL.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(" .")
    return s or "Unknown"


def client_initial(client_name: str) -> str:
    """
    "Patrick Gill" -> "P"
    If empty/unexpected -> "U"
    """
    name = (client_name or "").strip()
    if not name:
        return "U"
    return sanitize_component(name[0].upper())[:1] or "U"


def build_target_folder(
    *,
    root: str,
    client_name: str,
    year: str,
    property_address: str,
) -> str:
    """
    root is the fixed test root:
      /CostSeg Team Folder/Mark/Test Client Master

    Returns:
      <root>/Client Documents/<Initial>/<ClientName>/<Year>/MCSS/<Address>/EOB_Proposal_Payment
    """
    root = _norm_path(root)
    initial = client_initial(client_name)
    client = sanitize_component(client_name)
    yr = sanitize_component(year)
    addr = sanitize_component(property_address)

    return _norm_path(
        f"{root}"
        f"/Client Documents"
        f"/{initial}"
        f"/{client}"
        f"/{yr}"
        f"/MCSS"
        f"/{addr}"
        f"/EOB_Proposal_Payment"
    )


@dataclass
class DropboxClient:
    access_token: str
    timeout_s: int = 60
    _root_ns: Optional[str] = None

    def __post_init__(self) -> None:
        if not (self.access_token or "").strip():
            raise ValueError("Missing Dropbox access token")

    def _ensure_root_ns(self) -> Optional[str]:
        if self._root_ns is not None:
            return self._root_ns

        r = requests.post(
            f"{RPC}/users/get_current_account",
            headers={"Authorization": f"Bearer {self.access_token}"},
            timeout=30,
        )
        if r.status_code >= 400:
            raise DropboxError(f"get_current_account failed ({r.status_code}): {r.text[:500]}")

        data = r.json()
        # Log the Dropbox account identity for debugging
        email = data.get("email", "unknown")
        name = (data.get("name") or {}).get("display_name", "unknown")
        print(f"[DROPBOX_IDENTITY] email={email} name={name}")
        
        self._root_ns = (data.get("root_info") or {}).get("root_namespace_id")
        return self._root_ns

    def _headers(self, content: bool = False) -> dict[str, str]:
        headers = {"Authorization": f"Bearer {self.access_token}"}
        root_ns = self._ensure_root_ns()
        if root_ns:
            headers["Dropbox-API-Path-Root"] = json.dumps({".tag": "root", "root": root_ns})
        if not content:
            headers["Content-Type"] = "application/json"
        return headers

    def get_metadata(self, path: str) -> dict:
        path = _norm_path(path)
        r = requests.post(
            f"{RPC}/files/get_metadata",
            headers=self._headers(content=False),
            data=json.dumps({"path": path, "include_deleted": False}),
            timeout=self.timeout_s,
        )
        if r.status_code >= 400:
            raise DropboxError(f"get_metadata failed ({r.status_code}): {r.text[:500]}")
        return r.json()

    def create_folder(self, path: str) -> None:
        path = _norm_path(path)
        r = requests.post(
            f"{RPC}/files/create_folder_v2",
            headers=self._headers(content=False),
            data=json.dumps({"path": path, "autorename": False}),
            timeout=self.timeout_s,
        )

        if r.status_code == 409:
            try:
                md = self.get_metadata(path)
                if md.get(".tag") == "folder":
                    return
            except Exception:
                pass
            raise DropboxError(f"create_folder conflict at {path}: {r.text[:500]}")

        if r.status_code >= 400:
            raise DropboxError(f"create_folder failed ({r.status_code}): {r.text[:500]}")

    def ensure_parents(self, folder_path: str) -> None:
        folder_path = _norm_path(folder_path)
        parts = [p for p in folder_path.split("/") if p]
        cur = ""
        for part in parts:
            cur += "/" + part
            self.create_folder(cur)

    def upload_bytes(self, *, folder_path: str, filename: str, content: bytes, overwrite: bool = True) -> dict:
        folder_path = _norm_path(folder_path)
        filename = sanitize_component(filename)
        full_path = _norm_path(f"{folder_path}/{filename}")

        self.ensure_parents(folder_path)

        arg = {
            "path": full_path,
            "mode": "overwrite" if overwrite else "add",
            "autorename": False,
            "mute": False,
            "strict_conflict": False,
        }

        r = requests.post(
            f"{CONTENT}/files/upload",
            headers={
                **self._headers(content=True),
                "Content-Type": "application/octet-stream",
                "Dropbox-API-Arg": json.dumps(arg),
            },
            data=content,
            timeout=self.timeout_s,
        )

        if r.status_code >= 400:
            raise DropboxError(f"upload failed ({r.status_code}): {r.text[:500]}")
        return r.json()


def upload_eob_workbook(
    *,
    file_bytes: bytes,
    filename: str,
    client_name: str,
    year: str,
    property_address: str,
    logger,
) -> Optional[str]:
    """
    Wrapper used by app.py.
    Returns the uploaded path_display if uploaded, else None.
    """
    if os.getenv("DROPBOX_ENABLE", "").strip().lower() not in ("1", "true", "yes"):
        logger.info("Dropbox disabled: set DROPBOX_ENABLE=1 to enable upload")
        return None

    token = os.getenv("DROPBOX_ACCESS_TOKEN", "").strip()
    if not token:
        logger.warning("Dropbox upload skipped: DROPBOX_ACCESS_TOKEN not set")
        return None

    allowed_root = _norm_path(os.getenv(
        "DROPBOX_ALLOWED_ROOT",
        "/CostSeg Team Folder/Mark/Test Client Master"
    ))

    root = allowed_root  # fixed to allowed test root
    target_folder = build_target_folder(
        root=root,
        client_name=client_name,
        year=year,
        property_address=property_address,
    )

    logger.info(f"Dropbox enabled. allowed_root={allowed_root}")
    logger.info(f"Computed Dropbox target_folder={target_folder}")

    # Allowlist guardrail
    if not (target_folder == allowed_root or target_folder.startswith(allowed_root + "/")):
        logger.warning(f"Refusing Dropbox upload outside allowed root. allowed={allowed_root} got={target_folder}")
        return None

    dbx = DropboxClient(access_token=token)
    meta = dbx.upload_bytes(folder_path=target_folder, filename=filename, content=file_bytes, overwrite=True)
    path = meta.get("path_display") or meta.get("path_lower")
    logger.info(f"Dropbox upload OK: {path}")
    return path