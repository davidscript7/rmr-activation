#!/usr/bin/env python3
"""
============================================================
  RMR ACTIVATION — STEP 3: CONTRACT DETAIL → OUTPUT
  Reads the SQ01 report, processes dashes in Legacy Contract
  Number, merges with IW75 data to obtain CFIN DATE, and
  generates the final formatted RMR activation file.
============================================================
"""

import os
import sys
import json
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────
#  CONFIGURATION — Update column names here if the
#  SQ01 exported Excel changes its structure
# ─────────────────────────────────────────────────────────
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER  = os.path.join(SCRIPT_DIR, "input")
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, "output")
INTER_STEP2   = os.path.join(INPUT_FOLDER, ".step2_data.json")

# CONTRACT DETAIL column names (from SQ01 export)
COL_LEGACY_NUM  = "Legacy contract number"       # May contain dashes
COL_LEGACY_LINE = "Legacy contract line number"  # Line number
COL_TRANS_ID    = "Transaction ID"               # Output column
COL_ITEM_NUM    = "Item number in Document"      # Output column
# ─────────────────────────────────────────────────────────


def banner(text):
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)


def find_contract_file():
    """Finds the most recent CONTRACT DETAIL file in the INPUT folder."""
    files = sorted(
        [f for f in os.listdir(INPUT_FOLDER)
         if f.upper().startswith("CONTRACT DETAIL") and f.endswith(".xlsx")],
        reverse=True
    )
    if not files:
        print("  ❌ CONTRACT DETAIL file not found in the INPUT folder.")
        print(f"     Path: {INPUT_FOLDER}")
        print("     File name must start with: CONTRACT DETAIL")
        print("     Example: CONTRACT DETAIL 4.21.2026.xlsx")
        sys.exit(1)
    print(f"  📄 File: {files[0]}")
    return os.path.join(INPUT_FOLDER, files[0])


def detect_header_row(filepath, key_col):
    """Detects the real header row in SAP exports (skips title rows)."""
    import pandas as pd
    df_raw = pd.read_excel(filepath, header=None, nrows=30, dtype=str)
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values if str(v).strip() not in ("nan", "")]
        if key_col in vals:
            return i
    return 0


def read_contract_detail(filepath):
    """Reads the CONTRACT DETAIL file with automatic header detection."""
    import pandas as pd

    header_row = detect_header_row(filepath, COL_LEGACY_NUM)
    print(f"  📊 Header detected at row {header_row + 1}")

    df = pd.read_excel(filepath, header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    required = [COL_LEGACY_NUM, COL_LEGACY_LINE, COL_TRANS_ID, COL_ITEM_NUM]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        print(f"\n  ❌ Missing columns: {missing}")
        print(f"  Available columns: {list(df.columns)}")
        print("  👉 Update column names in the CONFIGURATION section of STEP3.py")
        sys.exit(1)

    for col in required:
        df[col] = df[col].str.strip()

    df = df[
        df[COL_LEGACY_NUM].notna() &
        (df[COL_LEGACY_NUM] != "") &
        (df[COL_LEGACY_NUM].str.lower() != "nan")
    ].copy()

    print(f"  Valid rows in CONTRACT DETAIL: {len(df)}")
    return df


def expand_dashes(df):
    """
    Expands rows where Legacy contract number contains dashes.
    Example: '123-456' with line '10' → two rows: ('123','10') and ('456','10')
    If the line number also contains dashes, they are paired positionally.
    """
    rows_out = []
    expanded = 0

    for _, row in df.iterrows():
        num_val  = str(row[COL_LEGACY_NUM]).strip()
        line_val = str(row[COL_LEGACY_LINE]).strip()

        if "-" in num_val:
            parts_num  = [p.strip() for p in num_val.split("-")]
            parts_line = [p.strip() for p in line_val.split("-")] if "-" in line_val else None

            for i, part in enumerate(parts_num):
                new_row = row.copy()
                new_row[COL_LEGACY_NUM] = part
                if parts_line and i < len(parts_line):
                    new_row[COL_LEGACY_LINE] = parts_line[i]
                rows_out.append(new_row)
            expanded += 1
        else:
            rows_out.append(row)

    import pandas as pd
    df_out = pd.DataFrame(rows_out).reset_index(drop=True)
    if expanded:
        print(f"  🔀 Rows expanded due to dashes: {expanded}")
        print(f"  Total rows after expansion: {len(df_out)}")
    return df_out


def to_num_str(v):
    """Converts a value to a numeric string without leading zeros."""
    try:
        return str(int(float(str(v).strip())))
    except (ValueError, TypeError):
        return str(v).strip()


def build_keys(df):
    """Builds composite key: Legacy contract number + Legacy contract line number."""
    df[COL_LEGACY_NUM]  = df[COL_LEGACY_NUM].apply(to_num_str)
    df[COL_LEGACY_LINE] = df[COL_LEGACY_LINE].apply(to_num_str)
    df["_KEY"] = df[COL_LEGACY_NUM] + "_" + df[COL_LEGACY_LINE]
    return df


def load_step2_data():
    """Loads the key → CFIN DATE mapping generated by Step 2."""
    if not os.path.exists(INTER_STEP2):
        print("\n  ⚠️  Step 2 data file not found (.step2_data.json)")
        print("     Make sure you ran STEP2.bat before this step.")
        sys.exit(1)
    with open(INTER_STEP2, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("key_to_cfin", {})


def merge_cfin(df, key_to_cfin):
    """Merges composite keys with IW75 data to obtain CFIN DATE."""
    def get_cfin(key):
        return key_to_cfin.get(str(key).strip())

    df["CFIN DATE"] = df["_KEY"].apply(get_cfin)
    matched = df["CFIN DATE"].notna().sum()
    total   = len(df)
    print(f"  ✅ CFIN DATE merged: {matched}/{total} rows")
    if matched < total:
        unmatched = df[df["CFIN DATE"].isna()][COL_LEGACY_NUM].head(5).tolist()
        print(f"  ⚠️  No match found for (first 5): {unmatched}")
    return df


def generate_output(df):
    """Generates the final formatted Excel file with 3 required columns."""
    import pandas as pd
    from datetime import datetime
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    output = df[[COL_TRANS_ID, COL_ITEM_NUM, "CFIN DATE"]].copy()
    output = output.dropna(how="all")
    output["CFIN DATE"] = pd.to_datetime(output["CFIN DATE"], errors="coerce").dt.strftime("%m/%d/%Y")

    today    = datetime.now().strftime("%m.%d.%Y")
    out_name = f"RMR_{today}.xlsx"
    out_path = os.path.join(OUTPUT_FOLDER, out_name)

    output.to_excel(out_path, index=False, sheet_name="RMR Activation")

    # Apply formatting
    wb = load_workbook(out_path)
    ws = wb.active

    header_fill  = PatternFill("solid", fgColor="1F4E79")
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    border_side  = Side(style="thin", color="CCCCCC")
    cell_border  = Border(left=border_side, right=border_side,
                          top=border_side, bottom=border_side)
    center_align = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 4):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.border    = cell_border
        cell.alignment = center_align

    for row_idx in range(2, ws.max_row + 1):
        fill_color = "EBF3FB" if row_idx % 2 == 0 else "FFFFFF"
        row_fill   = PatternFill("solid", fgColor=fill_color)
        for col_idx in range(1, 4):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill      = row_fill
            cell.border    = cell_border
            cell.alignment = center_align

    col_widths = [20, 25, 20]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    summary_row = ws.max_row + 2
    ws.cell(row=summary_row, column=1, value=f"Total records: {len(output)}")
    ws.cell(row=summary_row, column=1).font = Font(bold=True, color="1F4E79")
    ws.cell(row=summary_row + 1, column=1,
            value=f"Generated: {datetime.now().strftime('%m/%d/%Y %H:%M')}")

    wb.save(out_path)
    return out_path, len(output)


def main():
    banner("RMR ACTIVATION — STEP 3: CONTRACT DETAIL")

    print("🔧 Searching for CONTRACT DETAIL file...")
    filepath = find_contract_file()

    print("\n🔄 Reading CONTRACT DETAIL...")
    df = read_contract_detail(filepath)

    print("\n🔀 Processing dashes in Legacy Contract Number...")
    df = expand_dashes(df)

    print("\n🔑 Building merge keys...")
    df = build_keys(df)

    print("\n🔗 Loading IW75 data (STEP2)...")
    key_to_cfin = load_step2_data()

    print("\n📌 Merging with CFIN DATE...")
    df = merge_cfin(df, key_to_cfin)

    print("\n📝 Generating output file...")
    out_path, total = generate_output(df)

    banner(f"✅ PROCESS COMPLETED — {total} records")
    print(f"  Generated file: {os.path.basename(out_path)}")
    print(f"  Location: {out_path}")
    print("=" * 60)

    print("\n  📂 Opening file automatically...")
    try:
        os.startfile(out_path)
    except AttributeError:
        import subprocess
        subprocess.Popen(["xdg-open", out_path])
    except Exception as e:
        print(f"  ⚠️  Could not open automatically: {e}")
        print(f"  Open manually: {out_path}")

    print("\n  🎉 RMR Activation completed for today!")
    input("\n  Press ENTER to close...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress ENTER to close...")
        sys.exit(1)
