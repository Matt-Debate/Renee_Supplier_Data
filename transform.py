#!/usr/bin/env python3
"""
Renee(A)->Renee(B) converter

What it does
- Reads Renee(A).xlsx and extracts ONLY the 3 hockey-stick blocks:
    Block 1: B–F  (Model, Blade, Flex, Left, Right)
    Block 2: H–L
    Block 3: N–R
  Data is expected to start at row 5 with headers on row 4 (as in your files).

- Handles merged cells in Model/Blade by using the merged range's top-left value,
  then fill-downs Model/Blade while scanning.

- Defect rule (default ON):
    If a Left/Right quantity cell contains any non-ASCII characters (commonly Chinese
    like 瑕疵), that cell is treated as excluded (written as blank/None).

- Uses Renee(B).xlsx as a fixed template listing the rows/order of output.
  It fills Left/Right in that template by matching (Model, Blade, Flex).

- Optionally fills down missing Model/Blade cells in the B template (default ON).

IMPORTANT RELIABILITY CHANGE
- The B template may contain existing data. To prevent "data leakage", this script
  clears Left/Right columns for all rows before writing new values.

Outputs
- A generated .xlsx preserving Renee(B) formatting/styles.
- Optionally a CSV diff report comparing generated output to the provided B.

Usage
  python transform.py --a "Renee(A).xlsx" --b "Renee(B).xlsx" --out "Renee(B)_generated.xlsx"

Optional
  --no-defect-exclusion
  --no-filldown-b-template
  --diff-csv "diff_report.csv"
"""

import argparse
import math
import re
from collections import defaultdict
from typing import Any, Dict, Tuple, Optional, List

from openpyxl import load_workbook


# -----------------------------
# Parsing helpers
# -----------------------------

def has_non_ascii(s: str) -> bool:
    return any(ord(ch) > 127 for ch in s)

def norm_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None

def norm_model(v: Any) -> Optional[str]:
    return norm_text(v)

def norm_blade(v: Any) -> Optional[str]:
    s = norm_text(v)
    return s.upper() if s else None

def parse_flex(v: Any) -> Optional[int]:
    if v is None:
        return None
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        return int(round(v))
    if isinstance(v, str):
        m = re.search(r"\d+", v)
        return int(m.group(0)) if m else None
    return None

def parse_qty(v: Any, defect_exclusion: bool = True) -> Optional[int]:
    """
    Quantity parsing.
    - numeric -> int
    - string with digits -> int(digits)
    - if defect_exclusion and string contains non-ascii -> excluded (None)
    """
    if v is None:
        return None
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        if math.isfinite(v) and abs(v - round(v)) < 1e-9:
            return int(round(v))
        return int(v)
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        if defect_exclusion and has_non_ascii(s):
            return None
        m = re.search(r"\d+", s)
        return int(m.group(0)) if m else None
    return None


# -----------------------------
# Excel merged-cell helper
# -----------------------------

def merged_top_left_value(ws, cell) -> Any:
    """
    If cell is inside a merged range, return value of top-left cell of that range.
    Otherwise return cell.value.
    """
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            return ws.cell(rng.min_row, rng.min_col).value
    return cell.value


# -----------------------------
# Extract inventory from A
# -----------------------------

def build_inventory_from_a(path_a: str, defect_exclusion: bool = True) -> Dict[Tuple[str, str, int], Tuple[Optional[int], Optional[int]]]:
    wb_a = load_workbook(path_a, data_only=True)
    ws_a = wb_a.active

    # Fixed stick blocks
    blocks = [
        {"model_col": 2, "blade_col": 3, "flex_col": 4, "left_col": 5, "right_col": 6},       # B-F
        {"model_col": 8, "blade_col": 9, "flex_col": 10, "left_col": 11, "right_col": 12},    # H-L
        {"model_col": 14, "blade_col": 15, "flex_col": 16, "left_col": 17, "right_col": 18},  # N-R
    ]

    summed = defaultdict(lambda: {"L": 0, "R": 0})

    for blk in blocks:
        current_model = None
        current_blade = None

        for r in range(5, ws_a.max_row + 1):
            model_v = merged_top_left_value(ws_a, ws_a.cell(r, blk["model_col"]))
            blade_v = merged_top_left_value(ws_a, ws_a.cell(r, blk["blade_col"]))
            flex_v  = ws_a.cell(r, blk["flex_col"]).value
            left_v  = ws_a.cell(r, blk["left_col"]).value
            right_v = ws_a.cell(r, blk["right_col"]).value

            model_here = norm_model(model_v)
            blade_here = norm_blade(blade_v)

            if model_here:
                current_model = model_here
            if blade_here:
                current_blade = blade_here

            model = model_here or current_model
            blade = blade_here or current_blade
            flex = parse_flex(flex_v)

            if model is None or blade is None or flex is None:
                continue

            L = parse_qty(left_v, defect_exclusion=defect_exclusion)
            R = parse_qty(right_v, defect_exclusion=defect_exclusion)

            key = (model, blade, flex)
            if L is not None:
                summed[key]["L"] += L
            if R is not None:
                summed[key]["R"] += R

    # Convert 0 totals to None (blank)
    inv: Dict[Tuple[str, str, int], Tuple[Optional[int], Optional[int]]] = {}
    for k, v in summed.items():
        L = v["L"] if v["L"] > 0 else None
        R = v["R"] if v["R"] > 0 else None
        inv[k] = (L, R)

    return inv


# -----------------------------
# Apply inventory to B template
# -----------------------------

def apply_to_b_template(path_b: str,
                        out_path: str,
                        inv: Dict[Tuple[str, str, int], Tuple[Optional[int], Optional[int]]],
                        filldown_b_template: bool = True) -> None:
    """
    Writes a new file at out_path, preserving B formatting/styles.
    Assumes B uses columns:
        B: Model, C: Blade, D: Flex, E: Left, F: Right
    Header at row 1; data from row 2 down.

    Reliability:
    - Clears E/F for all rows before filling so existing template data never leaks through.
    """
    wb_out = load_workbook(path_b)  # keep styles
    ws = wb_out.active

    # Find last non-empty row in B for columns B-F
    def row_has_any(r: int) -> bool:
        for c in (2, 3, 4, 5, 6):
            v = ws.cell(r, c).value
            if v not in (None, ""):
                return True
        return False

    last_row = ws.max_row
    while last_row > 1 and not row_has_any(last_row):
        last_row -= 1

    # ---- CLEAR PASS (prevents data leakage from templates that already contain quantities) ----
    for r in range(2, last_row + 1):
        ws.cell(r, 5).value = None  # Left
        ws.cell(r, 6).value = None  # Right

    current_model = None
    current_blade = None

    for r in range(2, last_row + 1):
        model_cell = ws.cell(r, 2)
        blade_cell = ws.cell(r, 3)
        flex_cell  = ws.cell(r, 4)
        left_cell  = ws.cell(r, 5)
        right_cell = ws.cell(r, 6)

        model_here = norm_model(model_cell.value)
        blade_here = norm_blade(blade_cell.value)
        flex = parse_flex(flex_cell.value)

        if model_here:
            current_model = model_here
        if blade_here:
            current_blade = blade_here

        model = model_here or (current_model if filldown_b_template else None)
        blade = blade_here or (current_blade if filldown_b_template else None)

        # Optional: fix blank model/blade cells by filling them in
        if filldown_b_template:
            if (model_cell.value is None or str(model_cell.value).strip() == "") and model:
                model_cell.value = model
            if (blade_cell.value is None or str(blade_cell.value).strip() == "") and blade:
                blade_cell.value = blade

        if model is None or blade is None or flex is None:
            continue

        key = (model, blade, flex)
        newL, newR = inv.get(key, (None, None))

        left_cell.value = newL
        right_cell.value = newR

    wb_out.save(out_path)


# -----------------------------
# Diff report (optional)
# -----------------------------

def write_diff_csv(path_b_original: str, path_b_generated: str, diff_csv: str) -> None:
    """
    Compares values in B-F between original and generated and writes a CSV of cell diffs.
    """
    import csv

    wb_o = load_workbook(path_b_original, data_only=True)
    wb_g = load_workbook(path_b_generated, data_only=True)
    ws_o = wb_o.active
    ws_g = wb_g.active

    # Find last row based on original
    last_row = ws_o.max_row

    def row_has_any(ws, r: int) -> bool:
        return any(ws.cell(r, c).value not in (None, "") for c in (2, 3, 4, 5, 6))

    while last_row > 1 and not row_has_any(ws_o, last_row):
        last_row -= 1

    cols = [(2, "Model"), (3, "Blade"), (4, "Flex"), (5, "Left"), (6, "Right")]
    diffs: List[Dict[str, Any]] = []

    for r in range(1, last_row + 1):
        for c, name in cols:
            o = ws_o.cell(r, c).value
            g = ws_g.cell(r, c).value

            # Treat None and "" as equivalent blanks
            o_blank = (o is None) or (isinstance(o, str) and o.strip() == "")
            g_blank = (g is None) or (isinstance(g, str) and g.strip() == "")
            if o_blank and g_blank:
                continue

            if o != g:
                diffs.append({"row": r, "col": name, "orig_B": o, "generated": g})

    with open(diff_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["row", "col", "orig_B", "generated"])
        w.writeheader()
        w.writerows(diffs)


# -----------------------------
# Public API for Streamlit
# -----------------------------

def transform_files(path_a: str,
                    path_b: str,
                    out_path: str,
                    defect_exclusion: bool = True,
                    filldown_b_template: bool = True,
                    diff_csv: str | None = None) -> None:
    inv = build_inventory_from_a(path_a, defect_exclusion=defect_exclusion)
    apply_to_b_template(path_b, out_path, inv, filldown_b_template=filldown_b_template)
    if diff_csv:
        write_diff_csv(path_b, out_path, diff_csv)


# -----------------------------
# CLI
# -----------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--a", required=True, help="Path to Renee(A).xlsx")
    ap.add_argument("--b", required=True, help="Path to Renee(B).xlsx (template)")
    ap.add_argument("--out", required=True, help="Output path for generated B")
    ap.add_argument("--no-defect-exclusion", action="store_true",
                    help="If set, quantities with Chinese/non-ascii annotations will NOT be excluded.")
    ap.add_argument("--no-filldown-b-template", action="store_true",
                    help="If set, do NOT fill down blank Model/Blade cells inside B template.")
    ap.add_argument("--diff-csv", default=None,
                    help="Optional path to write a CSV diff report comparing provided B vs generated output.")

    args = ap.parse_args()

    transform_files(
        path_a=args.a,
        path_b=args.b,
        out_path=args.out,
        defect_exclusion=(not args.no_defect_exclusion),
        filldown_b_template=(not args.no_filldown_b_template),
        diff_csv=args.diff_csv,
    )


if __name__ == "__main__":
    main()