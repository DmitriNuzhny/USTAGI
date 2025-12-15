#!/usr/bin/env python3
"""eob_tool.io

JSON-first input loading for EOB calculators.

Why JSON:
- Monday.com payloads will arrive as structured key/value data.
- We want deterministic parsing (no Excel runtime dependency).

Supported input structures (all equivalent):

1) Cell-keyed (closest to legacy spreadsheets)
   {
     "B31": "SFR$$",
     "B1": 1620,
     "B2": 0.26,
     "B6": "100%",
     "B12": "700,000",
     "B32": "2021-01-01",
     "B34": 2025
   }

2) Field-keyed (human readable, mirrors the old text inputs)
   {
     "Tier": "SFR$$",
     "Interior SF": 1620,
     "Site Acres": 0.26,
     "Flooring (Sans Tile) %": "100%",
     "Basis": 700000,
     "Date Placed in Service": "2021-01-01",
     "Study Tax Year": 2025
   }

3) Wrapped:
   { "inputs": { ... } }  or  { "cells": { ... } }

We accept strings with:
- commas: "2,750,000"
- percents: "60%" (converted to 0.60)
- currency symbols: "$700,000"
- blank/None: treated as missing

This module returns a dict keyed by spreadsheet cell names (e.g., "B1")
so the calculator modules can stay spec-aligned.
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Optional


# --- Field -> cell mappings (based on legacy input conventions) ---

RESIDENTIAL_FIELD_TO_CELL: Dict[str, str] = {
    # Tier (not present in the old row list; it was the first line)
    "tier": "B31",
    "property tier": "B31",

    # Core geometry / counts
    "interior sf": "B1",
    "site acres": "B2",
    "bed cnt": "B3",
    "bath cnt": "B4",
    "tenant cnt": "B5",

    # Site fractions
    "flooring (sans tile) %": "B6",
    "flooring (sans tile)": "B6",
    "landscape %": "B7",
    "hardscape %": "B8",
    "parking lot %": "B9",

    # Equipment / basis / assumptions
    "solar cnt": "B10",
    "basis": "B12",
    "national avg $/sf (res.)": "B28",
    "national avg $/sf (res)": "B28",

    # Dates / tax year
    "date placed in service": "B32",
    "in-service date": "B32",
    "study tax year": "B34",
}

COMMERCIAL_FIELD_TO_CELL: Dict[str, str] = {
    "basis": "B1",
    "property type": "B2",
    "in-service date": "B32",
    "in service date": "B32",
    "date placed in service": "B32",
    "study tax year": "B34",
}


_CELL_RE = re.compile(r"^[A-Z]{1,3}\d{1,5}$")


def _norm_key(k: str) -> str:
    return re.sub(r"\s+", " ", str(k).strip()).lower()


def _parse_scalar(v: Any) -> Any:
    """Parse common spreadsheet-like strings into python scalars."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, (date, datetime)):
        return v.isoformat()

    s = str(v).strip()
    if s == "":
        return None

    # Percent -> fraction
    if s.endswith("%"):
        num = s[:-1].strip()
        num = num.replace(",", "").replace("$", "")
        try:
            return float(num) / 100.0
        except ValueError:
            return v  # leave as-is

    # Currency / comma number -> float
    cleaned = s.replace(",", "").replace("$", "")
    # Allow parentheses negatives: (1234) -> -1234
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1].strip()

    # Plain number?
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", cleaned):
        try:
            return float(cleaned)
        except ValueError:
            return v

    return v


def _extract_payload(d: Dict[str, Any]) -> Dict[str, Any]:
    if "inputs" in d and isinstance(d["inputs"], dict):
        return d["inputs"]
    if "cells" in d and isinstance(d["cells"], dict):
        return d["cells"]
    return d


def _field_to_cell(mode: str, field: str) -> Optional[str]:
    nk = _norm_key(field)
    if mode == "residential":
        return RESIDENTIAL_FIELD_TO_CELL.get(nk)
    if mode == "commercial":
        return COMMERCIAL_FIELD_TO_CELL.get(nk)
    return None


def load_inputs_from_json(mode: str, path: str | Path) -> Dict[str, Any]:
    """Load inputs from a JSON file and return a cell-keyed dict."""
    mode = str(mode).strip().lower()
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("JSON input must be an object/dict.")

    payload = _extract_payload(raw)
    if not isinstance(payload, dict):
        raise ValueError("JSON input must contain a dict under 'inputs' or 'cells', or be a dict itself.")

    cells: Dict[str, Any] = {}
    for k, v in payload.items():
        if k is None:
            continue
        ks = str(k).strip()
        if _CELL_RE.match(ks):
            cells[ks] = _parse_scalar(v)
            continue

        cell = _field_to_cell(mode, ks)
        if cell:
            cells[cell] = _parse_scalar(v)
        else:
            # Keep unknown fields (useful for debugging / forward compat)
            cells[ks] = _parse_scalar(v)

    return cells


def load_inputs_from_legacy_text(mode: str, path: str | Path) -> Dict[str, Any]:
    """Load the old text-file input format you used previously.

    Residential example:
      SFR$$
      Interior SF B1, 1620
      Flooring (Sans Tile) % B6, 100%
      ...

    Commercial example:
      Basis B1, 2,750,000
      Property Type B2, Medical Center
      ...
    """
    mode = str(mode).strip().lower()
    lines = [ln.strip() for ln in Path(path).read_text(encoding="utf-8").splitlines() if ln.strip()]
    if not lines:
        return {}

    cells: Dict[str, Any] = {}

    # Residential tier is often on first line by itself (e.g., SFR$$)
    if mode == "residential" and "," not in lines[0] and "B" not in lines[0]:
        tier = lines[0].strip()
        cells["B31"] = tier
        lines = lines[1:]

    for ln in lines:
        # Pattern: <Field> <Cell>, <Value>
        m = re.match(r"^(.*)\s+([A-Z]{1,3}\d{1,5})\s*,\s*(.*)$", ln)
        if m:
            field, cell, value = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
            cells[cell] = _parse_scalar(value)
            # also accept field-key if it maps and cell missing
            mapped = _field_to_cell(mode, field)
            if mapped and mapped not in cells:
                cells[mapped] = cells[cell]
            continue

        # Pattern: <Field>, <Value>
        m2 = re.match(r"^(.*)\s*,\s*(.*)$", ln)
        if m2:
            field, value = m2.group(1).strip(), m2.group(2).strip()
            mapped = _field_to_cell(mode, field)
            if mapped:
                cells[mapped] = _parse_scalar(value)
            else:
                cells[field] = _parse_scalar(value)

    return cells


def load_inputs(mode: str, input_path: Optional[str | Path]) -> Dict[str, Any]:
    """Convenience loader that picks the parser by extension.

    - .json => JSON loader (recommended)
    - .txt  => legacy text loader (kept for smooth transition)
    - None  => empty dict (calculators will fill defaults)
    """
    if not input_path:
        return {}

    p = Path(input_path)
    ext = p.suffix.lower()
    if ext == ".json":
        return load_inputs_from_json(mode, p)
    if ext == ".txt":
        return load_inputs_from_legacy_text(mode, p)

    raise ValueError(f"Unsupported input format: {p.name}. Use .json (preferred) or .txt (legacy).")


def example_json_schema(mode: str) -> Dict[str, Any]:
    """Return a minimal example JSON payload for a given mode."""
    mode = str(mode).strip().lower()
    if mode == "residential":
        return {
            "Tier": "SFR$$",
            "Interior SF": 1620,
            "Site Acres": 0.26,
            "Bed Cnt": 4,
            "Bath Cnt": 2,
            "Tenant Cnt": 1,
            "Flooring (Sans Tile) %": "100%",  # or 1.0
            "Landscape %": "60%",              # or 0.60
            "Hardscape %": "10%",              # or 0.10
            "Parking Lot %": "0%",
            "Solar Cnt": 0,
            "Basis": 700000,
            "National Avg $/SF (Res.)": 130,
            "Date Placed in Service": "2021-01-01",
            "Study Tax Year": 2025,
        }
    if mode == "commercial":
        return {
            "Basis": "2,750,000",
            "Property Type": "Medical Center",
            "In-Service Date": "2018-06-15",
            "Study Tax Year": 2025,
        }
    return {}
