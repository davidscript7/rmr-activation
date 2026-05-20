#!/usr/bin/env python3
"""
============================================================
RMR ACTIVATION — STEP 3: CONTRACT DETAIL → Excel
Writes items to RMR_Master.xlsx "Work" sheet.
Balanced assignment: each Transaction ID goes entirely to
ONE person (Hector or Daniel), balancing total items.
============================================================
"""

import os
import sys
import json
import warnings
from datetime import datetime
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────

SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER  = os.path.join(SCRIPT_DIR, "input")
INTER_STEP2   = os.path.join(INPUT_FOLDER, ".step2_data.json")
CONFIG_PATH   = os.path.join(SCRIPT_DIR, "config.json")

# CONTRACT DETAIL column names
COL_LEGACY_NUM  = "Legacy contract number"
COL_LEGACY_LINE = "Legacy contract Line number"
COL_TRANS_ID    = "Transaction ID"
COL_ITEM_NUM    = "Item Number in Document"

def banner(text):
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)

def load_config():
    if not os.path.exists(CONFIG_PATH):
        print("  ❌ config.json not found.")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def find_contract_file():
    files = sorted(
        [f for f in os.listdir(INPUT_FOLDER)
         if f.upper().startswith("CONTRACT DETAIL") and f.endswith(".xlsx")],
        reverse=True
    )
    if not files:
        print("  ❌ CONTRACT DETAIL file not found.")
        sys.exit(1)
    print(f"  📄 File: {files[0]}")
    return os.path.join(INPUT_FOLDER, files[0])

def detect_header_row(filepath, key_col):
    import pandas as pd
    df_raw = pd.read_excel(filepath, header=None, nrows=30, dtype=str)
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values if str(v).strip() not in ("nan", "")]
        if key_col in vals:
            return i
    return 0

def read_contract_detail(filepath):
    import pandas as pd
    header_row = detect_header_row(filepath, COL_LEGACY_NUM)
    print(f"  📊 Header detected at row {header_row + 1}")
    df = pd.read_excel(filepath, header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    required = [COL_LEGACY_NUM, COL_LEGACY_LINE, COL_TRANS_ID, COL_ITEM_NUM]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"\n  ❌ Missing columns: {missing}")
        sys.exit(1)
    for col in required:
        df[col] = df[col].str.strip()
    df = df[
        df[COL_LEGACY_NUM].notna() &
        (df[COL_LEGACY_NUM] != "") &
        (df[COL_LEGACY_NUM].str.lower() != "nan")
    ].copy()
    print(f"  Valid rows: {len(df)}")
    return df

def expand_dashes(df):
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
        print(f"  🔀 Rows expanded: {expanded}")
    return df_out

def to_num_str(v):
    try:
        return str(int(float(str(v).strip())))
    except (ValueError, TypeError):
        return str(v).strip()

def build_keys(df):
    df[COL_LEGACY_NUM]  = df[COL_LEGACY_NUM].apply(to_num_str)
    df[COL_LEGACY_LINE] = df[COL_LEGACY_LINE].apply(to_num_str)
    df["_KEY"] = df[COL_LEGACY_NUM] + "_" + df[COL_LEGACY_LINE]
    return df

def load_step2_data():
    if not os.path.exists(INTER_STEP2):
        print("\n  ⚠️  STEP2 data not found.")
        sys.exit(1)
    with open(INTER_STEP2, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("key_to_cfin", {}), data.get("date_label", "")

def merge_cfin(df, key_to_cfin):
    def get_cfin(key):
        return key_to_cfin.get(str(key).strip())
    df["CFIN DATE"] = df["_KEY"].apply(get_cfin)
    matched = df["CFIN DATE"].notna().sum()
    print(f"  ✅ CFIN DATE merged: {matched}/{len(df)} rows")
    return df

def assign_owners(df):
    import pandas as pd
    
    # Prepare output dataframe
    output = df[[COL_TRANS_ID, COL_ITEM_NUM, "CFIN DATE"]].copy()
    output = output.rename(columns={
        COL_TRANS_ID: "Transaction ID",
        COL_ITEM_NUM: "Item Number",
    })
    output = output.dropna(subset=["Transaction ID"])
    output["Project Number"] = df[COL_LEGACY_NUM].copy()
    
    # ── BALANCED ASSIGNMENT BY TRANSACTION ID ──
    # Each Transaction ID goes entirely to ONE person
    # Balance the total number of items between both
    
    # 1. Count items per Transaction ID
    tx_item_count = output.groupby("Transaction ID").size().to_dict()
    
    # 2. Sort Transaction IDs by item count (descending helps balance)
    sorted_txs = sorted(tx_item_count.items(), key=lambda x: x[1], reverse=True)
    
    # 3. Greedy assignment: assign each TX to person with fewer items
    hector_items = 0
    daniel_items = 0
    tx_to_owner = {}
    
    for tx_id, item_count in sorted_txs:
        if hector_items <= daniel_items:
            tx_to_owner[tx_id] = "Hector"
            hector_items += item_count
        else:
            tx_to_owner[tx_id] = "Daniel"
            daniel_items += item_count
    
    # 4. Assign owner to each row based on its Transaction ID
    output["Owner"] = output["Transaction ID"].map(tx_to_owner)
    
    # Add status columns
    output["Processed"] = "NO"
    output["Processing Date"] = None
    output["Status"] = "Pending sync"
    output["Id"] = None
    
    # Summary
    h_count = (output["Owner"] == "Hector").sum()
    d_count = (output["Owner"] == "Daniel").sum()
    h_txs = len([tx for tx, owner in tx_to_owner.items() if owner == "Hector"])
    d_txs = len([tx for tx, owner in tx_to_owner.items() if owner == "Daniel"])
    diff = abs(h_count - d_count)
    
    print(f"  👥 Balanced assignment:")
    print(f"     Hector: {h_count} items across {h_txs} Transaction IDs")
    print(f"     Daniel: {d_count} items across {d_txs} Transaction IDs")
    print(f"     Difference: {diff} items")
    
    return output

def write_to_excel(items_df, config):
    """Write items to RMR_Master.xlsm Work sheet using Excel COM"""
    
    try:
        import xlwings as xw
    except ImportError:
        print("  📦 Installing xlwings...")
        os.system("pip install xlwings --break-system-packages -q")
        import xlwings as xw
    
    rmr_master_path = config.get("rmr_master_path")
    
    if not rmr_master_path or not os.path.exists(rmr_master_path):
        print(f"  ❌ RMR_Master.xlsm not found")
        return 0
    
    print(f"  📂 Opening with Excel: {os.path.basename(rmr_master_path)}")
    
    try:
        # Open Excel (invisible)
        app = xw.App(visible=False)
        wb = app.books.open(rmr_master_path)
        
        # Get or create Work sheet
        if "Work" in [ws.name for ws in wb.sheets]:
            ws = wb.sheets["Work"]
        else:
            ws = wb.sheets.add("Work")
        
        # Clear existing data (keep row 1)
        ws.range("A2:Z10000").clear_contents()
        
        # Write headers
        headers = [
            "Transaction ID", "Item Number", "CFIN DATE", 
            "Project Number", "Owner", "Processed", 
            "Processing Date", "Status", "Id"
        ]
        ws.range("A1").value = headers
        
        # Prepare data as list of lists
        data = []
        for _, row in items_df.iterrows():
            # CFIN DATE formatting
            cfin_val = ""
            cfin_date = row.get("CFIN DATE")
            if cfin_date and str(cfin_date).lower() not in ("", "nan", "none"):
                try:
                    if isinstance(cfin_date, str):
                        dt = datetime.strptime(cfin_date, "%Y-%m-%d")
                        cfin_val = dt.strftime("%m/%d/%Y")
                    else:
                        cfin_val = cfin_date
                except:
                    cfin_val = str(cfin_date)
            
            data.append([
                str(row["Transaction ID"]),
                str(row["Item Number"]),
                cfin_val,
                str(row["Project Number"]),
                str(row["Owner"]),
                "NO",
                "",
                "Pending sync",
                ""
            ])
        
        # Write data starting at A2
        ws.range("A2").value = data
        
        # Format headers
        header_range = ws.range("A1:I1")
        header_range.color = (54, 96, 146)  # Blue
        header_range.api.Font.Color = 0xFFFFFF  # White text
        header_range.api.Font.Bold = True
        
        # Auto-fit columns
        ws.autofit(axis="columns")
        
        # Save and close
        wb.save()
        wb.close()
        app.quit()
        
        print(f"  ✅ Wrote {len(items_df)} items to 'Work' sheet")
        
        return len(items_df)
        
    except Exception as e:
        print(f"  ❌ Error: {e}")
        try:
            wb.close()
            app.quit()
        except:
            pass
        return 0


def main():
    banner("RMR ACTIVATION — STEP 3: CONTRACT DETAIL → Excel")
    
    print("\n⚙️  Loading configuration...")
    config = load_config()
    
    print("\n🔧 Searching for CONTRACT DETAIL...")
    filepath = find_contract_file()
    
    print("\n🔄 Reading CONTRACT DETAIL...")
    df = read_contract_detail(filepath)
    
    print("\n🔀 Processing dashes...")
    df = expand_dashes(df)
    
    print("\n🔑 Building keys...")
    df = build_keys(df)
    
    print("\n🔗 Loading STEP2 data...")
    key_to_cfin, date_label = load_step2_data()
    
    print("\n📌 Merging CFIN DATE...")
    df = merge_cfin(df, key_to_cfin)
    
    print("\n👥 Assigning owners (balanced by Transaction ID)...")
    items = assign_owners(df)
    
    print("\n📝 Writing to Excel...")
    added = write_to_excel(items, config)
    
    if added > 0:
        banner(f"✅ COMPLETED — {added} items added")
        h_count = (items["Owner"] == "Hector").sum()
        d_count = (items["Owner"] == "Daniel").sum()
        h_txs = len(items[items["Owner"] == "Hector"]["Transaction ID"].unique())
        d_txs = len(items[items["Owner"] == "Daniel"]["Transaction ID"].unique())
        diff = abs(h_count - d_count)
        print(f"  Hector: {h_count} items ({h_txs} Transaction IDs)")
        print(f"  Daniel: {d_count} items ({d_txs} Transaction IDs)")
        print(f"  Balance difference: {diff} items")
        print("=" * 60)
        
        print("\n  Next steps:")
        print("  1. Open RMR_Master.xlsx")
        print("  2. Review items in 'Work' sheet")
        print("  3. Run STEP4.bat to sync to SharePoint")
    else:
        print("\n  ❌ No items were added")
    
    input("\n  Press ENTER to close...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCancelled.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress ENTER...")
        sys.exit(1)