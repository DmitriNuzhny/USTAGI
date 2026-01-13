#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from openpyxl import load_workbook

# Adjust these cell addresses once you confirm them in your template.
# These are typical based on your earlier inspect output referencing 'Client Summary'!Gxx.
CLIENT_SUMMARY_INPUT_CELLS = {
    "Property Address": ("Client Summary", "G12"),
    "Building Use": ("Client Summary", "G20"),      # sometimes G13/G20 depending on template
    "Date Placed in Service": ("Client Summary", "G14"),
    "Basis": ("Client Summary", "G18"),
    "Tier": ("Client Summary", "G21"),
    "Accumulated Depreciation": ("Client Summary", "F26"),  # if used; else remove
    # Add more if needed
}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--address", default="")
    ap.add_argument("--building-use", default="SFR")
    ap.add_argument("--date", required=True)  # YYYY-MM-DD
    ap.add_argument("--basis", type=float, required=True)
    ap.add_argument("--tier", default="SFR$$")
    ap.add_argument("--accum-depr", type=float, default=0.0)
    args = ap.parse_args()

    wb = load_workbook(args.template, data_only=False)

    def put(key, value):
        sheet, cell = CLIENT_SUMMARY_INPUT_CELLS[key]
        wb[sheet][cell].value = value

    put("Property Address", args.address)
    put("Building Use", args.building_use)
    put("Date Placed in Service", args.date)
    put("Basis", args.basis)
    put("Tier", args.tier)
    put("Accumulated Depreciation", args.accum_depr)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote template-with-inputs to: {out}")
    print("Now open it in Excel, let it calculate, then Save As a new file (e.g. *_calced.xlsx).")

if __name__ == "__main__":
    main()