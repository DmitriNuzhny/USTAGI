from __future__ import annotations

from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Dict, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


RES_TEMPLATE = Path("templates") / "estimator_residential.xlsx"
COM_TEMPLATE = Path("templates") / "estimator_commercial.xlsx"


# These come directly from your debug_template output.
RES_MAP = {
    "property_address": "G9",
    "building_use": "G10",
    "date_in_service": "F11",
    "cost_basis": "G12",
    # Land Allocation was weird (value cell guess D13 is a formula). We'll support both:
    "land_allocation_text": "D13",   # e.g., "Per Depreciation Schedule"
    "land_allocation_amount": "G13", # if you want the $ value
    "building_basis": "G14",
    "improvements_included": "G15",
    "basis_for_cost_seg": "G16",
    "total_accelerated": "G19",
    "tax_savings_total": "G20",
    "estimated_addl_depr": "G22",
    # The “tax savings” for the additional depreciation is usually the next line.
    # If it’s not, adjust this to match the template.
    "tax_savings_addl": "G23",

    "table": {
        "start_row": 27,
        "n_years": 29,          # 2021–2049 (adjust if your template has more/less)
        "col_year": 2,          # B
        "col_5yr": 3,           # C
        "col_7yr": 4,           # D
        "col_15yr": 5,          # E
        "col_long": 6,          # F
        "col_with_css": 7,      # G
        "col_without_css": 8,   # H
    },
}

COM_MAP = {
    "property_address": "G9",
    "building_use": "G10",
    "date_in_service": "F11",
    "cost_basis": "G12",
    "land_allocation_text": "D13",
    "land_allocation_amount": "G13",
    "building_basis": "G14",
    "improvements_included": "G15",
    "basis_for_cost_seg": "G16",
    "total_accelerated": "G19",
    "tax_savings_total": "G20",
    "estimated_addl_depr": "G22",
    "tax_savings_addl": "G23",

    "table": {
        "start_row": 27,
        "n_years": 29,          # 2021–2049 (adjust if your template has more/less)
        "col_year": 2,          # B
        "col_5yr": 3,           # C
        "col_7yr": 4,           # D
        "col_15yr": 5,          # E
        "col_long": 6,          # F
        "col_with_css": 7,      # G
        "col_without_css": 8,   # H
    },
}


def _as_dict(obj: Any) -> Dict[str, Any]:
    if isinstance(obj, dict):
        return obj
    if is_dataclass(obj):
        return asdict(obj)
    if hasattr(obj, "__dict__"):
        return dict(obj.__dict__)
    raise TypeError(f"Unsupported result type: {type(obj)}")


def _ws_set(ws: Worksheet, addr: str, value: Any) -> None:
    # Overwrite formulas/#REF! with a concrete value.
    ws[addr].value = value


def _fill_table(ws: Worksheet, cfg: Dict[str, Any], yearly: Dict[int, Dict[str, Any]]) -> None:
    start_row = int(cfg["start_row"])      # 27
    n_years = int(cfg["n_years"])          # how many rows to fill (e.g., 29 for 2021-2049)
    start_year = int(cfg["start_year"])    # e.g., YEAR(date placed in service)

    col_year = int(cfg["col_year"])        # B=2
    c5 = int(cfg["col_5yr"])               # C=3
    c7 = int(cfg["col_7yr"])               # D=4
    c15 = int(cfg["col_15yr"])             # E=5
    clong = int(cfg["col_long"])           # F=6
    cwith = int(cfg["col_with_css"])       # G=7
    cwithout = int(cfg["col_without_css"]) # H=8

    # Write literal years + values row-by-row
    for i in range(n_years):
        year = start_year + i
        r = start_row + i

        ws.cell(r, col_year).value = year

        row = yearly.get(year, {})
        ws.cell(r, c5).value = row.get("5yr", 0)
        ws.cell(r, c7).value = row.get("7yr", 0)
        ws.cell(r, c15).value = row.get("15yr", 0)
        ws.cell(r, clong).value = row.get("long", 0)
        ws.cell(r, cwith).value = row.get("with_css", 0)
        ws.cell(r, cwithout).value = row.get("without_css", 0)



def _require_payload_shape(res: Dict[str, Any], mode: str) -> None:
    """
    Make it impossible to silently output a blank template.
    We require:
      res["summary"] : dict
      res["yearly"]  : dict[int -> dict]
    """
    if "summary" not in res or "yearly" not in res:
        keys = sorted(res.keys())
        raise ValueError(
            f"{mode} compute result does not include required keys 'summary' and 'yearly'.\n"
            f"Top-level keys present: {keys}\n\n"
            "Fix: add an adapter at the end of compute_residential/compute_commercial to return:\n"
            "  {'summary': {...}, 'yearly': {year: {'5yr':..,'7yr':..,'15yr':..,'long':..,'with_css':..,'without_css':..}}}\n"
        )


def _fill_estimator(ws: Worksheet, mapping: Dict[str, Any], result_obj: Any, mode: str) -> None:
    # If sheet is protected without a password, this allows writing.
    try:
        ws.protection.sheet = False
    except Exception:
        pass

    res = _as_dict(result_obj)
    _require_payload_shape(res, mode)

    summary: Dict[str, Any] = res["summary"] or {}
    yearly: Dict[int, Dict[str, Any]] = res["yearly"] or {}

    _ws_set(ws, mapping["property_address"], summary.get("property_address"))
    _ws_set(ws, mapping["building_use"], summary.get("building_use"))
    _ws_set(ws, mapping["date_in_service"], summary.get("date_placed_in_service"))
    _ws_set(ws, mapping["cost_basis"], summary.get("cost_basis"))

    # Land allocation: allow either text, amount, or both
    if "land_allocation_text" in mapping and summary.get("land_allocation_text") is not None:
        _ws_set(ws, mapping["land_allocation_text"], summary.get("land_allocation_text"))
    if "land_allocation_amount" in mapping and summary.get("land_allocation_amount") is not None:
        _ws_set(ws, mapping["land_allocation_amount"], summary.get("land_allocation_amount"))

    _ws_set(ws, mapping["building_basis"], summary.get("building_basis"))
    _ws_set(ws, mapping["improvements_included"], summary.get("improvements_included"))
    _ws_set(ws, mapping["basis_for_cost_seg"], summary.get("basis_for_cost_segregation"))

    _ws_set(ws, mapping["total_accelerated"], summary.get("total_accelerated"))
    _ws_set(ws, mapping["tax_savings_total"], summary.get("tax_savings_40pct_total_accel"))
    _ws_set(ws, mapping["estimated_addl_depr"], summary.get("estimated_additional_depr"))
    _ws_set(ws, mapping["tax_savings_addl"], summary.get("tax_savings_40pct_addl_depr"))

    # Set start_year for table
    dps = summary.get("date_placed_in_service")
    if hasattr(dps, "year"):
        start_year = int(dps.year)
    else:
        # ISO date string "2021-01-01" -> 2021
        start_year = int(str(dps)[:4]) if dps else min(yearly.keys()) if yearly else 2021

    # Set n_years to fill up to 2052
    last_year = 2052
    n_years = last_year - start_year + 1
    mapping["table"]["start_year"] = start_year
    mapping["table"]["n_years"] = n_years

    if yearly:
        _fill_table(ws, mapping["table"], yearly)


def write_residential_workbook(result: Any, out_path: Path | str) -> None:
    out_path = Path(out_path)
    wb = load_workbook(RES_TEMPLATE)
    ws = wb.active  # single sheet
    _fill_estimator(ws, RES_MAP, result, mode="residential")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def write_commercial_workbook(result: Any, out_path: Path | str) -> None:
    out_path = Path(out_path)
    wb = load_workbook(COM_TEMPLATE)
    ws = wb.active  # single sheet
    _fill_estimator(ws, COM_MAP, result, mode="commercial")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
