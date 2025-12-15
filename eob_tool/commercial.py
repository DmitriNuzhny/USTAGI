#!/usr/bin/env python3
"""
Commercial EOB calculator (spec-driven) with optional lookback.

Uses guideline fractions by property type (from the defaults workbook),
then allocates basis to 39/15/7/5 with Excel rounding and exact-basis adjustment.

Lookback:
- 39-year building (SL mid-month)
- 15-year (150% DB HY) + bonus
- 7-year (200% DB HY) + bonus
- 5-year (200% DB HY) + bonus
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Dict, Optional, Tuple

import pandas as pd

from .common import clamp01, excel_round, parse_date, compute_lookback, LookbackResult


DEFAULT_INPUTS: Dict[str, Any] = {
    "B1": 1000000.0,  # Basis
    "B2": "Bank",     # Property type
    "B32": None,      # In service
    "B34": None,      # Study year
}


@dataclass
class CommercialResult:
    inputs: Dict[str, Any]
    lookup_failed: bool
    matched_property_type: Optional[str]
    dep_life: Optional[Any]
    p39: float; p15: float; p7: float; p5: float; p_accel: float
    A39: int; A15: int; A7: int; A5: int
    s39: float; s15: float; s7: float; s5: float; s_accel: float
    # Lookback
    lookback_active: bool
    in_service_date: Optional[date]
    study_year: Optional[int]
    years_in_service: int
    lb_39: Optional[LookbackResult]
    lb_15: Optional[LookbackResult]
    lb_7: Optional[LookbackResult]
    lb_5: Optional[LookbackResult]
    total_current_year_depr: int
    total_cumulative_depr: int
    prior_years_depr: int


def load_guidelines(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _match_row(df: pd.DataFrame, prop_type: str) -> Tuple[Optional[pd.Series], bool]:
    """
    Match:
    1) exact case-insensitive
    2) contains case-insensitive
    pick first in table order deterministically
    """
    col = None
    for c in df.columns:
        if str(c).strip().lower() in {"property type", "property type guideline"}:
            col = c
            break
    if col is None:
        raise ValueError("Guideline file missing Property Type column.")

    q = str(prop_type or "").strip().lower()
    if not q:
        return None, True

    exact = df[df[col].astype(str).str.lower() == q]
    if len(exact) >= 1:
        return exact.iloc[0], False

    contains = df[df[col].astype(str).str.lower().str.contains(q, na=False)]
    if len(contains) >= 1:
        return contains.iloc[0], False

    return None, True


def compute_commercial(inputs: Dict[str, Any], guidelines_df: pd.DataFrame, today: Optional[date] = None) -> CommercialResult:
    if today is None:
        today = date.today()

    B = {**DEFAULT_INPUTS, **(inputs or {})}
    basis = max(float(B.get("B1", 0.0)), 0.0)
    prop_type = str(B.get("B2", "")).strip()

    row, failed = _match_row(guidelines_df, prop_type)

    if failed or row is None:
        return CommercialResult(
            inputs=dict(B),
            lookup_failed=True,
            matched_property_type=None,
            dep_life=None,
            p39=0.0, p15=0.0, p7=0.0, p5=0.0, p_accel=0.0,
            A39=0, A15=0, A7=0, A5=0,
            s39=0.0, s15=0.0, s7=0.0, s5=0.0, s_accel=0.0,
            lookback_active=False,
            in_service_date=None,
            study_year=None,
            years_in_service=0,
            lb_39=None, lb_15=None, lb_7=None, lb_5=None,
            total_current_year_depr=0,
            total_cumulative_depr=0,
            prior_years_depr=0,
        )

    # Pull fractions (column names vary between files/spec)
    def get_frac(*names: str, default: float = 0.0) -> float:
        for n in names:
            for c in guidelines_df.columns:
                if str(c).strip().lower() == n.strip().lower():
                    try:
                        return float(row.get(c, default))
                    except Exception:
                        return default
        return default

    p39 = clamp01(get_frac("39-yr", "39 yr", "39"))
    p15 = clamp01(get_frac("15-yr", "15 yr", "15"))
    p7  = clamp01(get_frac("7-yr", "7 yr", "7"))
    p5  = clamp01(get_frac("5-yr", "5 yr", "5"))
    p_accel = clamp01(get_frac("Total Accelerated", "Total Accelerated %", default=(p5+p7+p15)))

    dep_life = None
    for c in guidelines_df.columns:
        if str(c).strip().lower() in {"dep. life", "dep life", "dep. life (yrs)", "dep. life (years)"}:
            dep_life = row.get(c)
            break

    # Amounts (Excel rounding) + residual adjustment to match basis exactly
    a5 = basis * p5
    a7 = basis * p7
    a15 = basis * p15
    a39 = basis * p39

    A5 = int(excel_round(a5, 0))
    A7 = int(excel_round(a7, 0))
    A15 = int(excel_round(a15, 0))
    A39 = int(excel_round(a39, 0))

    sum_total = A5 + A7 + A15 + A39
    diff = int(excel_round(basis - sum_total, 0))
    if diff != 0:
        A39 = int(A39 + diff)

    if basis > 0:
        s5 = A5 / basis
        s7 = A7 / basis
        s15 = A15 / basis
        s39 = A39 / basis
        s_accel = s5 + s7 + s15
    else:
        s5 = s7 = s15 = s39 = s_accel = 0.0

    # Lookback activation
    in_service = parse_date(B.get("B32"))
    study_year = None
    if B.get("B34") is not None and str(B.get("B34")).strip() != "":
        try:
            study_year = int(float(str(B.get("B34")).replace(",", "")))
        except Exception:
            study_year = None

    lookback_active = (in_service is not None) and (study_year is not None)

    if not lookback_active:
        years_in_service = 0
        lb39 = lb15 = lb7 = lb5 = None
        total_cur = total_cum = prior = 0
    else:
        if in_service is None and study_year is not None:
            in_service = date(int(study_year), 1, 1)
        if study_year is None and in_service is not None:
            study_year = today.year

        assert in_service is not None and study_year is not None

        years_in_service = max(0, study_year - in_service.year + 1)

        lb39 = compute_lookback(
            basis=A39, in_service_date=in_service, study_year=study_year,
            asset_kind="building", is_residential_building=False
        )
        lb15 = compute_lookback(
            basis=A15, in_service_date=in_service, study_year=study_year,
            asset_kind="15", is_residential_building=False
        )
        lb7 = compute_lookback(
            basis=A7, in_service_date=in_service, study_year=study_year,
            asset_kind="7", is_residential_building=False
        )
        lb5 = compute_lookback(
            basis=A5, in_service_date=in_service, study_year=study_year,
            asset_kind="5", is_residential_building=False
        )

        total_cur = int(lb39.current_year_depreciation + lb15.current_year_depreciation + lb7.current_year_depreciation + lb5.current_year_depreciation)
        total_cum = int(lb39.cumulative_depreciation + lb15.cumulative_depreciation + lb7.cumulative_depreciation + lb5.cumulative_depreciation)
        prior = int(total_cum - total_cur) if years_in_service > 1 else 0

    res = CommercialResult(
        inputs=dict(B),
        lookup_failed=False,
        matched_property_type=str(row.get("Property Type", row.get("Property Type Guideline", prop_type))),
        dep_life=dep_life,
        p39=float(p39), p15=float(p15), p7=float(p7), p5=float(p5), p_accel=float(p_accel),
        A39=int(A39), A15=int(A15), A7=int(A7), A5=int(A5),
        s39=float(s39), s15=float(s15), s7=float(s7), s5=float(s5), s_accel=float(s_accel),
        lookback_active=lookback_active,
        in_service_date=in_service,
        study_year=study_year,
        years_in_service=years_in_service,
        lb_39=lb39 if lookback_active else None,
        lb_15=lb15 if lookback_active else None,
        lb_7=lb7 if lookback_active else None,
        lb_5=lb5 if lookback_active else None,
        total_current_year_depr=total_cur if lookback_active else 0,
        total_cumulative_depr=total_cum if lookback_active else 0,
        prior_years_depr=prior if lookback_active else 0,
    )

    return commercial_to_estimator_payload(res)


def commercial_to_estimator_payload(res: "CommercialResult") -> dict:
    """
    Convert CommercialResult into the payload expected by excel_writer.
    """
    inputs = res.inputs or {}

    summary = {
        "property_address": inputs.get("Property Address") or inputs.get("property_address") or "",
        "building_use": res.matched_property_type or "Commercial",
        "date_placed_in_service": (
            res.in_service_date.isoformat() if hasattr(res.in_service_date, "isoformat") else str(res.in_service_date)
        ),
        "cost_basis": inputs.get("Basis") or inputs.get("basis") or getattr(res, "A39", None),
        "land_allocation_text": "Per Depreciation Schedule",
        "land_allocation_amount": None,
        "building_basis": getattr(res, "A39", None),
        "improvements_included": getattr(res, "A15", None) + getattr(res, "A7", None) + getattr(res, "A5", None),
        "basis_for_cost_segregation": getattr(res, "A39", None) + getattr(res, "A15", None) + getattr(res, "A7", None) + getattr(res, "A5", None),
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

    yearly = {}

    def add_series(key: str, series):
        if not series:
            return
        if hasattr(series, 'year_by_year') and isinstance(series.year_by_year, list):
            for row in series.year_by_year:
                if hasattr(row, 'calendar_year') and hasattr(row, 'depreciation'):
                    y = row.calendar_year
                    v = row.depreciation
                    yearly.setdefault(y, {})[key] = v

    add_series("long", res.lb_39)
    add_series("15yr", res.lb_15)
    add_series("7yr", res.lb_7)
    add_series("5yr", res.lb_5)

    for y, row in yearly.items():
        with_css = 0
        any_val = False
        for k in ("5yr", "7yr", "15yr", "long"):
            if k in row and row[k] is not None:
                with_css += float(row[k])
                any_val = True
        row["with_css"] = with_css if any_val else None
        row["without_css"] = row.get("long")

    return {"summary": summary, "yearly": yearly}
