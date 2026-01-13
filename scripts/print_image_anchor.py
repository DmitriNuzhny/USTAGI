from __future__ import annotations

from pathlib import Path
from openpyxl import load_workbook

EMU_PER_PIXEL = 9525

def emu_to_px(v: int) -> float:
    return v / EMU_PER_PIXEL

def main():
    template = Path("outputs/test_with_logo.xlsx")
  # adjust if needed
    wb = load_workbook(template)
    ws = wb.active  # or wb["27.5 Estimate"] if you prefer

    imgs = getattr(ws, "_images", [])
    if not imgs:
        print("No images found on this sheet.")
        return

    for i, img in enumerate(imgs, start=1):
        a = img.anchor
        print(f"\nImage #{i}")
        print("  anchor type:", type(a).__name__)

        # OneCellAnchor has: _from (marker) + ext (size)
        if hasattr(a, "_from"):
            m = a._from
            print(f"  from: col={m.col}, row={m.row}, colOff(emu)={m.colOff}, rowOff(emu)={m.rowOff}")
            print(f"        colOff(px)={emu_to_px(m.colOff):.2f}, rowOff(px)={emu_to_px(m.rowOff):.2f}")

        if hasattr(a, "ext") and a.ext is not None:
            print(f"  ext:  cx(emu)={a.ext.cx}, cy(emu)={a.ext.cy}")
            print(f"       cx(px)={emu_to_px(a.ext.cx):.2f}, cy(px)={emu_to_px(a.ext.cy):.2f}")

        # openpyxl Image object also has width/height sometimes (not always reliable)
        try:
            print(f"  img.width={img.width}, img.height={img.height}")
        except Exception:
            pass

if __name__ == "__main__":
    main()
