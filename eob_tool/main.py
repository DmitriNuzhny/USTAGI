#!/usr/bin/env python3
"""
EOB Excel Generator (Residential & Commercial)

JSON-first inputs (recommended for Monday.com integration).

Examples:
  # Residential with defaults (from spec-driven code defaults)
  python -m eob_tool.main --mode residential --output out.xlsx

  # Residential from JSON
  python -m eob_tool.main --mode residential --input inputs.json --output out.xlsx

  # Commercial from JSON + default settings workbook (optional, if you want to mirror a specific guideline file)
  python -m eob_tool.main --mode commercial --input inputs.json --output out.xlsx \
    --guidelines "Commerical Only Estimator Default Settings (1).xlsx"

Notes:
- --input supports .json (preferred) and .txt (legacy).
- If --input is omitted, defaults from the calculators are used.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd

from .io import load_inputs
from .residential import compute_residential
from .commercial import compute_commercial, load_guidelines
from .excel_writer import write_residential_workbook, write_commercial_workbook


def load_commercial_guidelines_df():
    from pathlib import Path
    import pandas as pd

    p = Path("templates") / "commercial_estimator_default_settings.xlsx"
    if not p.exists():
        raise FileNotFoundError(
            "Missing commercial guidelines file: "
            "templates/commercial_estimator_default_settings.xlsx"
        )
    return pd.read_excel(p)


def main() -> int:
    ap = argparse.ArgumentParser(prog="eob_tool")
    ap.add_argument("--mode", required=True, choices=["residential", "commercial"], help="Which EOB to run")
    ap.add_argument(
        "--input",
        default=None,
        help="Path to input file (.json preferred; .txt legacy). If omitted, calculator defaults are used.",
    )
    ap.add_argument("--output", required=True, help="Path to output .xlsx")
    ap.add_argument(
        "--guidelines",
        default=None,
        help="Commercial-only: optional path to a guideline/default-settings workbook used to seed some assumptions.",
    )

    args = ap.parse_args()
    mode = args.mode.lower()

    B = load_inputs(mode, args.input)

    out_path = Path(args.output)

    if mode == "residential":
        res = compute_residential(B)
        write_residential_workbook(res, out_path)
    else:
        guidelines_df = load_commercial_guidelines_df()
        com = compute_commercial(B, guidelines_df)
        write_commercial_workbook(com, out_path)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
