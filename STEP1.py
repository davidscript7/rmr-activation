#!/usr/bin/env python3
"""
============================================================
  RMR ACTIVATION — STEP 1: (Power Query version)
  
  Reads SharePoint data via Power Query, applies filters,
  and prepares Project Numbers for IW75.
  
 
============================================================
"""

import os
import sys
import json
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────
# CONFIGURATION — SharePoint List column names
# ─────────────────────────────────────────────────────────

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(SCRIPT_DIR, "input")
INTER_FILE   = os.path.join(INPUT_FOLDER, ".step1_data.json")

# Power Query file
QUERY_FILE = os.path.join(INPUT_FOLDER, "RMR_Master.xlsx")
QUERY_SHEET = "RMR Activation - To Process"  # Primera hoja

# Column names from SharePoint List
COL_PROJECT_NUM = "Project Number"
COL_CFIN_DATE   = "Actual Install Completion Date (Solomon)"
COL_INSTALL_ONLY = "InstallOnly"
COL_RRR         = "RRR"
COL_STATUS      = "Status"
COL_RMR_STATUS  = "RMR Status"

# Filter values
STATUS_VALID = "Invoiced/Install Processed"
RMR_STATUS_VALID = ["Not Yet Processed", ""]  # blank o "Not Yet Processed"

# ─────────────────────────────────────────────────────────

def banner(text):
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)


def check_query_file():
    """Verify Power Query file exists"""
    if not os.path.exists(QUERY_FILE):
        print("\n" + "=" * 60)
        print("  ⚠️  POWER QUERY FILE NOT FOUND")
        print("=" * 60)
        print()
        print(f"  Expected: {QUERY_FILE}")
        print()
        print("  📋 FIRST TIME SETUP:")
        print("  1. Create Excel with Power Query connection to SharePoint")
        print("  2. Save as: RMR_Master.xlsx")
        print(f"  3. Place in: {INPUT_FOLDER}")
        print()
        print("  📖 See POWER_QUERY_SETUP.md for detailed instructions")
        print("=" * 60)
        input("\n  Press ENTER to exit...")
        sys.exit(1)
    
    print(f"  ✅ Power Query file found: {os.path.basename(QUERY_FILE)}")
    
    size_kb = os.path.getsize(QUERY_FILE) // 1024
    print(f"     File size: {size_kb} KB")


def auto_refresh_query():
    """
    Automatically refresh Power Query data from SharePoint.
    This eliminates the manual step of opening Excel and clicking Refresh.
    """
    print("\n🔄 Auto-refreshing Power Query from SharePoint...")
    print("  ⏳ This may take 10-20 seconds...")
    
    try:
        import win32com.client
    except ImportError:
        print("  📦 Installing pywin32 (needed for auto-refresh)...")
        os.system("pip install pywin32 -q")
        try:
            import win32com.client
        except ImportError:
            print("  ⚠️  Could not install pywin32 - manual refresh required")
            print("     Open RMR_Master.xlsx → Ctrl+Alt+F5 before running")
            input("\n  Press ENTER to continue anyway...")
            return False
    
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(os.path.abspath(QUERY_FILE))
        
        # Refresh all queries
        wb.RefreshAll()
        
        # Wait for refresh to complete
        excel.CalculateUntilAsyncQueriesDone()
        
        # Save and close
        wb.Save()
        wb.Close(SaveChanges=False)
        excel.Quit()
        
        print("  ✅ Data refreshed from SharePoint successfully")
        return True
        
    except Exception as e:
        print(f"  ⚠️  Auto-refresh failed: {e}")
        print("     Continuing with existing data...")
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception:
            pass
        return False


def read_monitoring():
    """Read the Power Query Excel file"""
    import pandas as pd
    
    print("\n📊 Reading Monitoring data...")
    
    try:
        df = pd.read_excel(QUERY_FILE, sheet_name=QUERY_SHEET, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        print(f"  📋 Sheet: {QUERY_SHEET}")
        print(f"  📊 Rows read: {len(df)}")
        return df
        
    except Exception as e:
        print(f"\n  ❌ Error reading file: {e}")
        print(f"\n  Available sheets in file:")
        import pandas as pd
        xl_file = pd.ExcelFile(QUERY_FILE)
        for sheet in xl_file.sheet_names:
            print(f"     - {sheet}")
        print(f"\n  💡 Update QUERY_SHEET variable if sheet name is different")
        sys.exit(1)


def verify_columns(df):
    """Verify all required columns exist"""
    required = [
        COL_PROJECT_NUM,
        COL_CFIN_DATE,
        COL_INSTALL_ONLY,
        COL_RRR,
        COL_STATUS,
        COL_RMR_STATUS
    ]
    
    missing = [c for c in required if c not in df.columns]
    
    if missing:
        print(f"\n  ❌ Missing columns: {missing}")
        print(f"\n  Available columns:")
        for col in df.columns:
            print(f"     - {col}")
        print(f"\n  💡 Update column names in CONFIGURATION section of STEP1.py")
        sys.exit(1)
    
    print("  ✅ All required columns found")


def apply_filters(df):
    """Apply business logic filters"""
    import pandas as pd
    
    print("\n🔍 Applying filters...")
    
    initial_count = len(df)
    
    # Filter 1: Install Only = FALSE
    print("  📌 Filter 1: Install Only = FALSE")
    cond_install = df[COL_INSTALL_ONLY].astype(str).str.upper() == "FALSE"
    df = df[cond_install].copy()
    print(f"     Rows after filter: {len(df)}")
    
    # Filter 2: RRR = blank
    print("  📌 Filter 2: RRR = blank")
    cond_rrr = (
        df[COL_RRR].isna() | 
        (df[COL_RRR].astype(str).str.strip() == "") |
        (df[COL_RRR].astype(str).str.lower() == "nan")
    )
    df = df[cond_rrr].copy()
    print(f"     Rows after filter: {len(df)}")
    
    # Filter 3: Status = "Invoiced/Install Processed"
    print(f"  📌 Filter 3: Status = '{STATUS_VALID}'")
    cond_status = df[COL_STATUS].astype(str).str.strip() == STATUS_VALID
    df = df[cond_status].copy()
    print(f"     Rows after filter: {len(df)}")
    
    # Filter 4: RMR Status = "Not Yet Processed" OR blank
    print("  📌 Filter 4: RMR Status = 'Not Yet Processed' or blank")
    cond_rmr = (
        df[COL_RMR_STATUS].isna() |
        (df[COL_RMR_STATUS].astype(str).str.strip() == "") |
        (df[COL_RMR_STATUS].astype(str).str.lower() == "nan") |
        (df[COL_RMR_STATUS].astype(str).str.strip() == "Not Yet Processed")
    )
    df = df[cond_rmr].copy()
    print(f"     Rows after filter: {len(df)}")
    
    # Filter 5: CFIN DATE <= today or blank
    print("  📌 Filter 5: CFIN DATE <= today or blank")
    df["_cfin_dt"] = pd.to_datetime(df[COL_CFIN_DATE], errors="coerce")
    today = pd.Timestamp.today().normalize()
    cond_cfin = df["_cfin_dt"].isna() | (df["_cfin_dt"] <= today)
    df = df[cond_cfin].copy()
    print(f"     Rows after filter: {len(df)}")
    
    # Remove rows with empty Project Number
    print("  📌 Final: Remove empty Project Numbers")
    df[COL_PROJECT_NUM] = df[COL_PROJECT_NUM].astype(str).str.strip()
    df = df[
        df[COL_PROJECT_NUM].notna() &
        (df[COL_PROJECT_NUM] != "") &
        (df[COL_PROJECT_NUM].str.lower() != "nan")
    ].copy()
    print(f"     Rows after cleanup: {len(df)}")
    
    filtered_count = len(df)
    removed_count = initial_count - filtered_count
    
    print(f"\n  📊 Summary:")
    print(f"     Initial rows:   {initial_count}")
    print(f"     Filtered rows:  {filtered_count}")
    print(f"     Removed:        {removed_count}")
    
    if filtered_count == 0:
        print("\n  ⚠️  No records match the filters!")
        print("     Check if:")
        print("     - Data exists in SharePoint")
        print("     - Filter criteria are correct")
        print("     - Power Query refreshed successfully")
        sys.exit(1)
    
    return df


def save_intermediate(df):
    """Save data for STEP2"""
    import pandas as pd
    from datetime import datetime
    
    # Build mapping: Project Number → CFIN DATE
    mapping = {}
    for _, row in df.iterrows():
        proj_num = str(row[COL_PROJECT_NUM]).strip()
        cfin = row.get("_cfin_dt")
        
        if pd.notna(cfin):
            mapping[proj_num] = cfin.strftime("%Y-%m-%d")
        else:
            mapping[proj_num] = None
    
    # Use today's date for file naming
    date_label = datetime.now().strftime("%m.%d.%Y")
    
    data = {
        "sap_to_cfin": mapping,
        "date_label": date_label,
        "sheets_processed": [QUERY_SHEET]
    }
    
    os.makedirs(os.path.dirname(INTER_FILE), exist_ok=True)
    with open(INTER_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    
    print(f"\n  💾 Data saved for STEP2 ({len(mapping)} records)")
    print(f"  📅 Date label: {date_label}")


def copy_to_clipboard(text):
    """Copy text to clipboard"""
    try:
        import pyperclip
        pyperclip.copy(text)
        return True
    except Exception:
        return False


def main():
    banner("RMR ACTIVATION — STEP 1: MONITORING")
    
    print("\n🔧 Checking Power Query file...")
    check_query_file()
    
    # Auto-refresh Power Query data
    auto_refresh_query()
    
    # Read data
    df = read_monitoring()
    
    # Verify columns
    verify_columns(df)
    
    # Apply filters
    df_filtered = apply_filters(df)
    
    # Save intermediate data for STEP2
    save_intermediate(df_filtered)
    
    # Extract and format Project Numbers for SAP
    def format_for_sap(number):
        """Remove C0 prefix and add asterisks for SAP IW75"""
        number = str(number).strip()
        # Remove C0 prefix if exists
        if number.startswith("C0"):
            number = number[2:]  # Remove first 2 characters (C0)
        elif number.startswith("C"):
            number = number[1:]  # Remove first character only (C)
        return f"*{number}*"
    
    project_numbers = [format_for_sap(num) for num in df_filtered[COL_PROJECT_NUM]]
    project_block = "\n".join(project_numbers)
    
    banner(f"SAP NUMBERS READY (formatted) — {len(project_numbers)} records")
    print(project_block)
    print("=" * 60)
    print("  (Format: *number* without C0 prefix for SAP IW75)")
    print("=" * 60)
    
    if copy_to_clipboard(project_block):
        print("\n📋 Numbers copied to clipboard ✅")
    else:
        print("\n⚠️  Could not copy automatically. Please copy manually.")
    
    # Get date label for IW75 file naming
    from datetime import datetime
    date_label = datetime.now().strftime("%m.%d.%Y")
    month = datetime.now().strftime("%B")
    
    banner("NEXT STEPS — SAP IW75")
    print("  1. Open SAP Fiori → transaction IW75")
    print("  2. 'Your Reference' field → multiple selection button")
    print("  3. In the multiple selection window:")
    print("     → Click 'Clipboard' icon")
    print("     → Numbers will paste automatically")
    print("  4. Press F8 (Execute)")
    print("  5. Export to Excel:")
    print("     Menu → List → Save/Send → Local File → Spreadsheet")
    print(f"  6. Save in: {INPUT_FOLDER}")
    print(f"     Name: IW75 {month} {date_label}.xlsx")
    print()
    print("  7. Once you have the file → run STEP2.bat")
    print("=" * 60)
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