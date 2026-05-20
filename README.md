# RMR Activation — Process Automation

> **Automates the daily RMR Activation workflow** using Power Query, Python, and bidirectional SharePoint synchronization, eliminating manual filtering, data merging, and status tracking across SAP transactions IW75 and SQ01.

---

## Overview

This project automates a daily operational process that previously required extensive manual work across Excel, SAP, and SharePoint. The automation is divided into **4 sequential steps** with automatic data synchronization.

```
[STEP 1] Power Query auto-refresh → filter SharePoint List data
            ↓
        SAP IW75 — paste Project Numbers, export report
            ↓
[STEP 2] Process IW75 report, build keys, prepare SQ01 list
            ↓
        SAP SQ01 — paste Sales Documents, export report
            ↓
[STEP 3] Process Contract Detail, merge CFIN DATE, balanced assignment
            ↓
        Write to RMR_Master.xlsm "Work" sheet
            ↓
[STEP 4] Sync Work sheet → SharePoint Lists (bidirectional)
            ↓
        Auto-update Lista 1 when projects complete
```

---

## Key Features

### Architecture Highlights
- **Power Query integration** — Direct SharePoint connection eliminates authentication friction
- **Balanced workload distribution** — Transaction IDs assigned entirely to one person, minimizing total item difference
- **Bidirectional sync** — Excel ↔ SharePoint Lists with automatic status updates
- **Auto-completion detection** — Updates Lista 1 when all items in a project are processed

### Technical Improvements
- **No programmatic auth required** — Uses Office 365 credentials via Power Query
- **Excel COM automation** — Auto-refresh queries, write formatted data with xlwings
- **Smart prefix handling** — Automatically matches Project Numbers with/without C0 prefix
- **Dash expansion** — `123-456` in Legacy Contract Number splits into separate rows

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3 | Core automation logic |
| pandas | Data filtering and merging |
| openpyxl | Excel reading and header detection |
| xlwings | Excel COM automation (write/refresh) |
| pyperclip | Clipboard integration for SAP |
| pywin32 | Power Query auto-refresh |
| Office365-REST-Python-Client | SharePoint List API integration |

---

## Project Structure

```
rmr-activation/
│
├── STEP1.py                    # Power Query refresh, filter, extract Project Numbers
├── STEP2.py                    # IW75: process report, build keys, prepare SQ01 list
├── STEP3.py                    # Contract Detail: merge, balanced assignment, write to Excel
├── STEP4.py                    # Sync Excel → SharePoint Lists (bidirectional)
│
├── STEP1.bat                   # Double-click launcher for Step 1
├── STEP2.bat                   # Double-click launcher for Step 2
├── STEP3.bat                   # Double-click launcher for Step 3
├── STEP4.bat                   # Double-click launcher for Step 4
├── INSTALL_DEPENDENCIES.bat    # Run once to install Python packages
│
├── config.example.json         # Configuration template
│
├── input/                      # SAP exports + RMR_Master.xlsm
│   └── RMR_Master.xlsm        # Master file with Power Query connections
│
└── output/                     # Generated files (if needed)
```

---

## Setup (First Time Only)

### 1. Prerequisites

- Python 3.8+ installed with **"Add Python to PATH"** checked
- Windows OS
- Microsoft Excel (Office 365 or 2016+)
- Access to SAP Fiori (IW75, SQ01)
- SharePoint access to:
  - **Lista 1**: Install Projects to Process (1070)
  - **Lista 2**: RMR ACTIVATION

### 2. Install Python packages

Double-click `INSTALL_DEPENDENCIES.bat`. Wait for all `[OK]` confirmations.

### 3. Configure

```bash
# Copy the example config and fill in your values
copy config.example.json config.json
```

Open `config.json` with Notepad and update:
- `sharepoint_site_url` — Your SharePoint site URL
- `list_1_name` / `list_2_name` — SharePoint List names
- `rmr_master_path` — Path to RMR_Master.xlsm
- `base_folder` — Local path to this project folder
- `owner` — Your name (Hector or Daniel)

### 4. Create input folder

```
rmr-activation/
└── input/     ← create this folder
    └── RMR_Master.xlsm  ← place your master file here
```

### 5. Power Query Setup (RMR_Master.xlsm)

The RMR_Master.xlsm file must have two Power Query connections:

**Query 1: "RMR Activation - To Process"**
- Connects to SharePoint Lista 1: Install Projects to Process (1070)
- Filters: Status = "Invoiced/Install Processed", RMR Status = blank or "Not Yet Processed"
- Loads to first sheet

**Query 2: "RMR ACTIVATION"**
- Connects to SharePoint Lista 2: RMR ACTIVATION  
- Loads to second sheet (for STEP4 sync)

See Power Query connection wizard: Data → Get Data → From SharePoint Online List

---

## Daily Usage

| Step | Action |
|------|--------|
| `STEP1.bat` | Auto-refresh Power Query, filter data, copy Project Numbers to clipboard |
| **SAP IW75** | Paste numbers → run report → export Excel to `input/` |
| `STEP2.bat` | Process IW75, copy SQ01 list to clipboard |
| **SAP SQ01** | Paste list → run report → export Excel to `input/` |
| `STEP3.bat` | Process Contract Detail, merge CFIN DATE, write to Work sheet with balanced assignment |
| `STEP4.bat` | Sync Work sheet → SharePoint Lists, auto-update completed projects |

### File naming convention (required for auto-detection)

| File | Required format | Example |
|------|----------------|---------|
| IW75 export | `IW75 [Month] [Date].xlsx` | `IW75 May 5.20.2026.xlsx` |
| Contract Detail | `CONTRACT DETAIL [Date].xlsx` | `CONTRACT DETAIL 5.20.2026.xlsx` |
| Master file | `RMR_Master.xlsm` (in input/) | Fixed name |

---

## Balanced Assignment Algorithm

**How Transaction IDs are distributed between Hector and Daniel:**

1. **Group by Transaction ID** — Count items per Transaction ID
2. **Sort descending** — Largest Transaction IDs first (helps balance)
3. **Greedy assignment** — Assign each Transaction ID to person with fewer accumulated items

**Example:**
- TX 1000: 10 items → Hector (0 → 10)
- TX 1001: 8 items → Daniel (0 → 8)
- TX 1002: 5 items → Daniel (8 → 13)
- TX 1003: 2 items → Hector (10 → 12)

**Result:** Hector: 12 items (2 TXs), Daniel: 13 items (2 TXs), Difference: 1 item

**Guarantee:** Each Transaction ID is assigned **entirely** to one person (never split), while minimizing total item difference.

---

## SharePoint Integration

### Lista 1: Install Projects to Process (1070)
- **Source** for STEP1 (via Power Query)
- **Target** for STEP4 completion updates
- Updated when all items in a project are marked as Processed

### Lista 2: RMR ACTIVATION
- **Target** for STEP4 item updates
- Stores individual Transaction ID items with Owner, Processed status, Processing Date

### STEP4 Sync Flow

1. Read Work sheet → identify items pending sync
2. Update Lista 2 with Processed status, Processing Date, Project Number
3. Check Lista 2 → find projects where ALL items are Processed = YES
4. Update Lista 1 → set RMR Status = "RMR Processed", RMR Processed By = Owner
5. Mark Work sheet items as "Synced"

**Note:** STEP4 is currently in development and may require manual verification.

---

## Configuration Reference

| Key | Description |
|-----|-------------|
| `sharepoint_site_url` | SharePoint site URL |
| `list_1_name` | Lista 1 display name |
| `list_2_name` | Lista 2 display name |
| `rmr_master_path` | Full path to RMR_Master.xlsm |
| `base_folder` | Local project folder path |
| `owner` | Your name (for filtering/assignment) |
| `client_id` / `client_secret` | Optional (for API auth if needed) |

---

## Troubleshooting

### STEP1: Power Query won't refresh
- Open RMR_Master.xlsm manually
- Data → Refresh All (Ctrl+Alt+F5)
- Check SharePoint credentials (File → Account)

### STEP2: No match for Project Numbers
- Verify IW75 export has "Your Reference" column
- Check if numbers need C0 prefix removal
- STEP2 now tries: direct match, C0prefix, C prefix

### STEP3: Items not appearing in Work sheet
- Verify config.json has correct `rmr_master_path`
- Check if Excel is open (close it before running STEP3)
- Verify xlwings is installed

### STEP4: Sync errors
- Verify SharePoint permissions (edit access to both Lists)
- Check `client_id` / `client_secret` if using app-based auth
- Run with admin privileges if Windows Integrated Auth fails

---

## Author

**Héctor Gallego** — [@davidscript7](https://github.com/davidscript7)

---

*Built to reduce daily manual workload through Power Query, Python automation, and bidirectional SharePoint synchronization.*