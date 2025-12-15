#!/usr/bin/env python3
"""
Residential EOB calculator (spec-driven) with optional lookback.

Implements calculations from Residential_EOB_Dashboard_Spec.md:
- Excel-consistent rounding order and behavior.
- Tier logic (SFR$/$$/$$$ + MFR variants).
- Optional lookback (building 27.5-year SL mid-month; 5-year & 15-year with bonus).

This module returns structured results for Excel export.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Dict, Optional, Tuple

from .common import clamp01, excel_round, parse_date, compute_lookback, LookbackResult


TIER_RATES: Dict[str, Dict[str, float]] = {
    "SFR$":   {"rate17": 21638, "rate18": 257, "rate19": 283, "rate20": 4,  "rate21": 15300},
    "SFR$$":  {"rate17": 28565, "rate18": 257, "rate19": 283, "rate20": 11, "rate21": 15300},
    "SFR$$$": {"rate17": 32691, "rate18": 257, "rate19": 283, "rate20": 11, "rate21": 15300},
}
MFR_TIERS = {"MFR$", "MFR$$", "MFR$$$"}

DEFAULT_INPUTS: Dict[str, Any] = {
    "B1": 2000.0,   # Interior SF
    "B2": 0.22,     # Site Acres
    "B3": 3.0,      # Bed Cnt
    "B4": 2.0,      # Bath Cnt
    "B5": 1.0,      # Tenant Cnt
    "B6": 1.0,      # Flooring fraction
    "B7": 0.60,     # Landscape fraction
    "B8": 0.10,     # Hardscape fraction
    "B9": 0.00,     # Parking fraction
    "B10": 0.00,    # Solar Cnt
    "B11": 0.00,    # Pool Cnt (display only)
    "B12": 300000.0,# Basis for allocation
    "B28": 130.0,   # National Avg $/SF (Res.)
    "B31": "SFR$",  # Tier
    "B32": None,    # In-service date (optional)
    "B34": None,    # Study year (optional)
}


@dataclass
class ResidentialResult:
    inputs: Dict[str, Any]
    tier_display: str
    building_type: str  # SFR/MFR
    # Section 1
    C6_internal: int
    area_sf: int
    landscape_ac: float
    hardscape_ac: float
    parking_ac: float
    landscape_sf: int
    hardscape_sf: int
    parking_sf: int
    # Section 3/4/5
    D17: int; D18: int; D19: int; D20: int; D21: int; D22: int
    D23: int; D24: int; D25: int; D26: int; D27: int
    D28: int; D29: int; D30: int
    # Section 2 allocation
    B13: int; B14: int; B15: int; B16: int
    C13_pct: float; C14_pct: float; C15_pct: float; C16_pct: float
    D13: int; D14: int; D15: int; D16: int
    # Lookback (optional)
    lookback_active: bool
    in_service_date: Optional[date]
    study_year: Optional[int]
    years_in_service: int
    lookback_building: Optional[LookbackResult]
    lookback_5: Optional[LookbackResult]
    lookback_15: Optional[LookbackResult]
    total_current_year_depr: int
    total_cumulative_depr: int
    prior_years_depr: int


def _apply_tier_logic(B: Dict[str, Any]) -> Tuple[Dict[str, Any], str, str]:
    eff = dict(B)
    raw = str(eff.get("B31", "SFR$")).strip().upper()
    building_type = "SFR"
    tier_display = raw

    if raw in MFR_TIERS:
        building_type = "MFR"
        suffix = raw[3:]
        eff["B31"] = "SFR" + suffix
        eff["B28"] = 200.0
    elif raw not in TIER_RATES:
        eff["B31"] = "SFR$"
        tier_display = "SFR$"
    else:
        eff["B31"] = raw

    return eff, tier_display, building_type


def compute_residential(inputs: Dict[str, Any], today: Optional[date] = None) -> ResidentialResult:
    """
    Compute Residential EOB results per spec (Sections 1-5) + optional lookback.
    `inputs` should provide B-cells (B1.. etc). Missing keys use defaults.
    """
    if today is None:
        today = date.today()

    B0 = {**DEFAULT_INPUTS, **(inputs or {})}
    B, tier_display, building_type = _apply_tier_logic(B0)

    # --- normalize/clamp input fields ---
    B1 = float(B.get("B1", 0.0))
    B2 = float(B.get("B2", 0.0))
    B3 = float(B.get("B3", 0.0))
    B4 = float(B.get("B4", 0.0))
    B5 = float(B.get("B5", 0.0))
    B6 = clamp01(B.get("B6", 0.0))
    B7 = clamp01(B.get("B7", 0.0))
    B8 = clamp01(B.get("B8", 0.0))
    B9 = clamp01(B.get("B9", 0.0))
    B10 = float(B.get("B10", 0.0))
    B12 = float(B.get("B12", 0.0))
    B28 = float(B.get("B28", 0.0))

    # Section 1: Flooring helper (round after multiply)
    C6_internal = int(excel_round(B1 * B6, 0))

    # Section 1: site fractions w/ 70% cap
    s = B7 + B8 + B9
    if s <= 0.70 + 1e-12:
        p7, p8, p9 = B7, B8, B9
    else:
        f = 0.70 / s
        p7, p8, p9 = B7 * f, B8 * f, B9 * f

    area_raw = B2 * 43560.0
    area_sf = int(excel_round(area_raw, 0))

    landscape_sf = int(excel_round(p7 * area_sf, 0))
    hardscape_sf = int(excel_round(p8 * area_sf, 0))
    parking_sf = int(excel_round(p9 * area_sf, 0))

    landscape_ac = float(excel_round(p7 * B2, 2))
    hardscape_ac = float(excel_round(p8 * B2, 2))
    parking_ac = float(excel_round(p9 * B2, 2))

    # Section 3: 5-year property dollars (integer by construction)
    rates = TIER_RATES[str(B.get("B31", "SFR$"))]
    D17 = int(excel_round(rates["rate17"] * B5, 0))
    D18 = int(excel_round(rates["rate18"] * B3, 0))
    D19 = int(excel_round(rates["rate19"] * B4, 0))
    D20 = int(excel_round(rates["rate20"] * C6_internal, 0))
    D21 = int(excel_round(rates["rate21"] * B10, 0))
    D22 = int(D17 + D18 + D19 + D20 + D21)

    # Section 4: 15-year property dollars (use rounded SF first, then multiply, then ROUND)
    D23 = 0
    D24 = int(excel_round(2.786 * landscape_sf, 0))
    D25 = int(excel_round(8.0 * hardscape_sf, 0))
    D26 = int(excel_round(7.87 * parking_sf, 0))
    D27 = int(D24 + D25 + D26)

    # Section 5: Building
    D28 = int(excel_round(B28 * B1, 0))
    D29 = int(-D22)
    D30 = int(D28 + D29)

    # Section 2: EOB allocation
    B13 = int(D30)
    B14 = int(D22)
    B15 = int(D27)
    B16 = int(B13 + B14 + B15)

    if B16 > 0:
        C13 = B13 / B16
        C14 = B14 / B16
        C15 = B15 / B16
        C16 = C14 + C15
    else:
        C13 = C14 = C15 = C16 = 0.0

    C13_pct = float(excel_round(C13 * 100.0, 2))
    C14_pct = float(excel_round(C14 * 100.0, 2))
    C15_pct = float(excel_round(C15 * 100.0, 2))
    C16_pct = float(excel_round(C16 * 100.0, 2))

    D13 = int(excel_round(C13 * B12, 0))
    D14 = int(excel_round(C14 * B12, 0))
    D15 = int(excel_round(C15 * B12, 0))
    D16 = int(D14 + D15)

    # Lookback activation
    in_service = parse_date(B.get("B32"))
    study_year = None
    if B.get("B34") is not None and str(B.get("B34")).strip() != "":
        try:
            study_year = int(float(str(B.get("B34")).replace(",", "")))
        except Exception:
            study_year = None

    lookback_active = (in_service is not None) or (study_year is not None)

    if not lookback_active:
        years_in_service = 0
        lb_building = lb_5 = lb_15 = None
        total_cur = total_cum = prior = 0
    else:
        if in_service is None and study_year is not None:
            in_service = date(int(study_year), 1, 1)
        if study_year is None and in_service is not None:
            study_year = today.year

        assert in_service is not None and study_year is not None

        years_in_service = max(0, study_year - in_service.year + 1)

        lb_building = compute_lookback(
            basis=D13, in_service_date=in_service, study_year=study_year,
            asset_kind="building", is_residential_building=True
        )
        lb_5 = compute_lookback(
            basis=D14, in_service_date=in_service, study_year=study_year,
            asset_kind="5", is_residential_building=True
        )
        lb_15 = compute_lookback(
            basis=D15, in_service_date=in_service, study_year=study_year,
            asset_kind="15", is_residential_building=True
        )

        total_cur = int(lb_building.current_year_depreciation + lb_5.current_year_depreciation + lb_15.current_year_depreciation)
        total_cum = int(lb_building.cumulative_depreciation + lb_5.cumulative_depreciation + lb_15.cumulative_depreciation)
        prior = int(total_cum - total_cur) if years_in_service > 1 else 0

    res = ResidentialResult(
        inputs=dict(B),
        tier_display=tier_display,
        building_type=building_type,
        C6_internal=C6_internal,
        area_sf=area_sf,
        landscape_ac=landscape_ac,
        hardscape_ac=hardscape_ac,
        parking_ac=parking_ac,
        landscape_sf=landscape_sf,
        hardscape_sf=hardscape_sf,
        parking_sf=parking_sf,
        D17=D17, D18=D18, D19=D19, D20=D20, D21=D21, D22=D22,
        D23=D23, D24=D24, D25=D25, D26=D26, D27=D27,
        D28=D28, D29=D29, D30=D30,
        B13=B13, B14=B14, B15=B15, B16=B16,
        C13_pct=C13_pct, C14_pct=C14_pct, C15_pct=C15_pct, C16_pct=C16_pct,
        D13=D13, D14=D14, D15=D15, D16=D16,
        lookback_active=lookback_active,
        in_service_date=in_service,
        study_year=study_year,
        years_in_service=years_in_service,
        lookback_building=lb_building if lookback_active else None,
        lookback_5=lb_5 if lookback_active else None,
        lookback_15=lb_15 if lookback_active else None,
        total_current_year_depr=total_cur if lookback_active else 0,
        total_cumulative_depr=total_cum if lookback_active else 0,
        prior_years_depr=prior if lookback_active else 0,
    )

    return residential_to_estimator_payload(res)


def residential_to_estimator_payload(res: "ResidentialResult") -> dict:
    """
    Convert ResidentialResult (cell-style outputs) into the payload expected by excel_writer.
    Uses existing computed fields; no new math.
    """
    # Inputs may store address etc eventually; for now allow missing.
    inputs = res.inputs or {}

    # Header / summary block
    summary = {
        "property_address": inputs.get("Property Address") or inputs.get("property_address") or "",
        "building_use": res.building_type or "Single Family Residence",
        "date_placed_in_service": (
            res.in_service_date.isoformat() if hasattr(res.in_service_date, "isoformat") else str(res.in_service_date)
        ),
        "cost_basis": inputs.get("Basis") or inputs.get("basis") or getattr(res, "B13", None),
        # Your template has two land allocation cells (text + amount). Use text default.
        "land_allocation_text": "Per Depreciation Schedule",
        "land_allocation_amount": None,

        # Basis lines — map these after we confirm what B13..B16 represent in your model.
        # For now, use what we have (best-effort):
        "building_basis": getattr(res, "B14", None),
        "improvements_included": getattr(res, "B15", None),
        "basis_for_cost_segregation": getattr(res, "B16", None),

        # Totals block
        "total_accelerated": getattr(res, "total_current_year_depr", None),
        "tax_savings_40pct_total_accel": (
            getattr(res, "total_current_year_depr", 0) * 0.40
            if getattr(res, "total_current_year_depr", None) is not None
            else None
        ),
        "estimated_additional_depr": getattr(res, "prior_years_depr", None),
        "tax_savings_40pct_addl_depr": (
            getattr(res, "prior_years_depr", 0) * 0.40
            if getattr(res, "prior_years_depr", None) is not None
            else None
        ),
    }

    # Year table: use lookback_* objects if present.
    # Expected each lookback_* to behave like {year: amount} or list of rows.
    yearly = {}

    # Helper to add a series into yearly
    def add_series(key: str, series):
        if not series:
            return
        # If it's a LookbackResult dataclass
        if hasattr(series, 'year_by_year') and isinstance(series.year_by_year, list):
            for row in series.year_by_year:
                if hasattr(row, 'calendar_year') and hasattr(row, 'depreciation'):
                    y = row.calendar_year
                    v = row.depreciation
                    yearly.setdefault(y, {})[key] = v
        elif isinstance(series, dict):
            for y, v in series.items():
                try:
                    y = int(y)
                except Exception:
                    continue
                yearly.setdefault(y, {})[key] = v
        elif isinstance(series, list):
            # allow list of (year, value) tuples
            for row in series:
                if isinstance(row, (tuple, list)) and len(row) >= 2:
                    try:
                        y = int(row[0])
                    except Exception:
                        continue
                    yearly.setdefault(y, {})[key] = row[1]

    # These are your computed lookback schedules
    add_series("long", res.lookback_building)
    add_series("5yr", res.lookback_5)
    add_series("15yr", res.lookback_15)

    # If you don’t have 7-year in residential (often you do), leave it blank.
    # Totals columns:
    # with_css = sum of components
    for y, row in yearly.items():
        with_css = 0
        any_val = False
        for k in ("5yr", "7yr", "15yr", "long"):
            if k in row and row[k] is not None:
                with_css += float(row[k])
                any_val = True
        row["with_css"] = with_css if any_val else None

        # "without_css": building-only depreciation. Use long if present.
        row["without_css"] = row.get("long")

    return {"summary": summary, "yearly": yearly}
