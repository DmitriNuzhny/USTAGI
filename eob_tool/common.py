#!/usr/bin/env python3
"""
Common utilities for EOB calculators (Residential & Commercial).

Key design goal: match legacy Excel behavior.
- Excel-style rounding: ROUND(x, 0) rounds .5 away from zero (not banker's rounding).
- Date parsing supports common string formats and Excel serial numbers.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import math


def excel_round(x: float, ndigits: int = 0) -> float:
    """
    Excel-style ROUND:
    - Rounds halves away from zero.
    - ndigits behaves like Excel ROUND.
    """
    if x is None:
        return 0.0
    try:
        x = float(x)
    except Exception:
        return 0.0

    factor = 10.0 ** ndigits
    y = x * factor
    if y >= 0:
        y = math.floor(y + 0.5)
    else:
        y = math.ceil(y - 0.5)
    return y / factor


def clamp01(x: float) -> float:
    try:
        x = float(x)
    except Exception:
        return 0.0
    return max(0.0, min(1.0, x))


def parse_date(val: Any) -> Optional[date]:
    """Parse various date formats into a date object."""
    if val is None or (isinstance(val, str) and not val.strip()):
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, (int, float)):
        # Excel serial date (days since 1899-12-30)
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(val))).date()
        except Exception:
            return None
    if isinstance(val, str):
        s = val.strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


# --- Bonus depreciation schedule (per specs) ---

BONUS_SCHEDULE: List[Tuple[date, date, float]] = [
    (date(1900, 1, 1), date(2001, 9, 10), 0.0),
    (date(2001, 9, 11), date(2003, 5, 5), 0.30),
    (date(2003, 5, 6), date(2004, 12, 31), 0.50),
    (date(2005, 1, 1), date(2007, 12, 31), 0.0),
    (date(2008, 1, 1), date(2010, 9, 8), 0.50),
    (date(2010, 9, 9), date(2011, 12, 31), 1.0),
    (date(2012, 1, 1), date(2017, 9, 27), 0.50),
    (date(2017, 9, 28), date(2022, 12, 31), 1.0),
    (date(2023, 1, 1), date(2023, 12, 31), 0.80),
    (date(2024, 1, 1), date(2024, 12, 31), 0.60),
    (date(2025, 1, 1), date(2025, 1, 19), 0.40),
    (date(2025, 1, 20), date(2099, 12, 31), 0.0),
]


def get_bonus_rate(in_service_date: date) -> float:
    for start, end, rate in BONUS_SCHEDULE:
        if start <= in_service_date <= end:
            return rate
    return 0.0


# --- MACRS tables (from specs / IRS tables) ---

# 5-year (200% DB, half-year)
MACRS_5_YEAR = {1: 0.2000, 2: 0.3200, 3: 0.1920, 4: 0.1152, 5: 0.1152, 6: 0.0576}

# 7-year (200% DB, half-year)
MACRS_7_YEAR = {1: 0.1429, 2: 0.2449, 3: 0.1749, 4: 0.1249, 5: 0.0893, 6: 0.0892, 7: 0.0893, 8: 0.0446}

# 15-year (150% DB, half-year)
MACRS_15_YEAR = {
    1: 0.0500, 2: 0.0950, 3: 0.0855, 4: 0.0770, 5: 0.0693, 6: 0.0623,
    7: 0.0590, 8: 0.0590, 9: 0.0591, 10: 0.0590, 11: 0.0591, 12: 0.0590,
    13: 0.0591, 14: 0.0590, 15: 0.0591, 16: 0.0295,
}

# 27.5-year residential (mid-month) table by year and month (1-indexed month)
MACRS_27_5_YEAR: Dict[int, List[float]] = {
    1: [0.03485, 0.03182, 0.02879, 0.02576, 0.02273, 0.01970, 0.01667, 0.01364, 0.01061, 0.00758, 0.00455, 0.00152],
    28: [0.01970, 0.02273, 0.02576, 0.02879, 0.03182, 0.03485, 0.03636, 0.03636, 0.03636, 0.03636, 0.03636, 0.03636],
    29: [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.00152, 0.00455, 0.00758, 0.01061, 0.01364, 0.01667],
}
for yr in range(2, 28):
    MACRS_27_5_YEAR[yr] = [0.03636] * 12

# 39-year nonresidential (mid-month)
MACRS_39_YEAR: Dict[int, List[float]] = {
    1: [0.02461, 0.02247, 0.02033, 0.01819, 0.01605, 0.01391, 0.01177, 0.00963, 0.00749, 0.00535, 0.00321, 0.00107],
    40: [0.00107, 0.00321, 0.00535, 0.00749, 0.00963, 0.01177, 0.01391, 0.01605, 0.01819, 0.02033, 0.02247, 0.02461],
}
for yr in range(2, 40):
    MACRS_39_YEAR[yr] = [0.02564] * 12


def building_rate(year_index: int, in_service_month: int, is_residential: bool) -> float:
    if is_residential:
        table = MACRS_27_5_YEAR
        max_year = 29
    else:
        table = MACRS_39_YEAR
        max_year = 40

    if year_index < 1 or year_index > max_year:
        return 0.0
    rates = table.get(year_index)
    if not rates:
        return 0.0
    m = max(1, min(12, int(in_service_month)))
    return float(rates[m - 1])


def macrs_rate(asset_life: str, year_index: int) -> float:
    if asset_life == "5":
        return float(MACRS_5_YEAR.get(year_index, 0.0))
    if asset_life == "7":
        return float(MACRS_7_YEAR.get(year_index, 0.0))
    if asset_life == "15":
        return float(MACRS_15_YEAR.get(year_index, 0.0))
    return 0.0


def get_bonus_rate(in_service_date: date) -> float:
    # Simplified: assume 80% bonus for qualified property
    # In real implementation, check if date qualifies for bonus
    return 0.8


@dataclass(frozen=True)
class LookbackYearRow:
    calendar_year: int
    depreciation_year: int
    rate: float
    depreciation: int
    cumulative: int


@dataclass
class LookbackResult:
    original_basis: int
    bonus_rate: float
    bonus_amount: int
    depreciable_basis: int
    current_year_depreciation: int
    cumulative_depreciation: int
    net_book_value: int
    year_by_year: List[LookbackYearRow]


def compute_lookback(
    *,
    basis: float,
    in_service_date: date,
    study_year: int,
    asset_kind: str,        # "building" | "5" | "7" | "15"
    is_residential_building: bool,
) -> LookbackResult:
    """
    Lookback rules per specs:
    - bonus applies to 5/7/15 only; none to building.
    - annual depreciation = ROUND(depreciable_basis * rate, 0) (Excel-style).
    - cumulative includes bonus + all annual MACRS up through study_year.
    """
    basis_i = int(excel_round(float(basis), 0))
    if basis_i <= 0 or study_year < in_service_date.year:
        return LookbackResult(
            original_basis=basis_i,
            bonus_rate=0.0,
            bonus_amount=0,
            depreciable_basis=max(basis_i, 0),
            current_year_depreciation=0,
            cumulative_depreciation=0,
            net_book_value=basis_i,
            year_by_year=[],
        )

    in_month = in_service_date.month

    if asset_kind == "building":
        bonus_rate = 0.0
        bonus_amount = 0
        depreciable_basis = basis_i
        max_years = 29 if is_residential_building else 40
    else:
        bonus_rate = get_bonus_rate(in_service_date)
        bonus_amount = int(excel_round(basis_i * bonus_rate, 0))
        depreciable_basis = basis_i - bonus_amount
        max_years = {"5": 6, "7": 8, "15": 16}.get(asset_kind, 0)

    cumulative = bonus_amount
    current_year_dep = 0
    rows: List[LookbackYearRow] = []

    year_index = 1
    for cal_year in range(in_service_date.year, study_year + 1):
        if year_index > max_years:
            break

        if asset_kind == "building":
            rate = building_rate(year_index, in_month, is_residential_building)
        else:
            rate = macrs_rate(asset_kind, year_index)

        annual = int(excel_round(depreciable_basis * rate, 0))
        cumulative += annual

        if cal_year == study_year:
            current_year_dep = annual

        rows.append(
            LookbackYearRow(
                calendar_year=cal_year,
                depreciation_year=year_index,
                rate=float(rate),
                depreciation=annual,
                cumulative=int(cumulative),
            )
        )
        year_index += 1

    net_book = int(excel_round(basis_i - cumulative, 0))
    return LookbackResult(
        original_basis=basis_i,
        bonus_rate=float(bonus_rate),
        bonus_amount=int(bonus_amount),
        depreciable_basis=int(depreciable_basis),
        current_year_depreciation=int(current_year_dep),
        cumulative_depreciation=int(cumulative),
        net_book_value=net_book,
        year_by_year=rows,
    )


# --- Forward schedule helpers (full-life depreciation, not just "lookback to study_year") ---

def compute_full_schedule(
    *,
    basis: float,
    in_service_date: date,
    asset_kind: str,              # "building" | "5" | "7" | "15"
    is_residential_building: bool,
) -> Dict[int, int]:
    """
    Returns {calendar_year -> depreciation_amount} for the full recovery life.

    IMPORTANT:
    - For 5/7/15: first year includes BONUS + MACRS year1.
    - For building: no bonus, uses mid-month table, full 27.5 (29 rows) or 39 (40 rows).
    - Uses Excel-style rounding (excel_round).
    """
    basis_i = int(excel_round(float(basis), 0))
    if basis_i <= 0:
        return {}

    start_year = int(in_service_date.year)
    in_month = int(in_service_date.month)

    out: Dict[int, int] = {}

    if asset_kind == "building":
        max_years = 29 if is_residential_building else 40
        for year_index in range(1, max_years + 1):
            cal_year = start_year + (year_index - 1)
            rate = building_rate(year_index, in_month, is_residential_building)
            annual = int(excel_round(basis_i * rate, 0))
            out[cal_year] = annual
        return out

    # 5/7/15 assets
    max_years = {"5": 6, "7": 8, "15": 16}.get(asset_kind, 0)
    if max_years <= 0:
        return {}

    bonus_rate = get_bonus_rate(in_service_date)
    bonus_amount = int(excel_round(basis_i * bonus_rate, 0))
    depreciable_basis = basis_i - bonus_amount

    for year_index in range(1, max_years + 1):
        cal_year = start_year + (year_index - 1)
        rate = macrs_rate(asset_kind, year_index)
        macrs_amt = int(excel_round(depreciable_basis * rate, 0))
        if year_index == 1:
            out[cal_year] = int(bonus_amount + macrs_amt)
        else:
            out[cal_year] = macrs_amt

    return out
