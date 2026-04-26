# RMR Activation — Process Automation

> **Automates the daily RMR Activation workflow**, eliminating manual filtering, data merging, and formatting across SAP transactions IW75 and SQ01.

---

## Overview

This project was built to automate a daily operational process that previously required manual work in Excel and SAP. The automation is divided into **3 sequential steps**, with pauses between each for SAP interaction.

```
[STEP 1] Download & filter Monitoring file (SharePoint)
            ↓
        SAP IW75 — paste SAP numbers, export report
            ↓
[STEP 2] Process IW75 report, build keys, prepare SQ01 list
            ↓
        SAP SQ01 — paste Sales Documents, export report
            ↓
[STEP 3] Process Contract Detail, merge CFIN DATE, generate final file
            ↓
        OUTPUT: Transaction ID | Item number in Document | CFIN DATE
```

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3 | Core automation logic |
| pandas | Data filtering and merging |
| openpyxl | Excel reading and formatting |
| msal | Microsoft 365 authentication (SharePoint) |
| requests | SharePoint file download |
| pyperclip | Clipboard integration |
| Windows BAT | One-click script execution for non-technical users |

---

## Project Structure

```
rmr-activation/
│
├── STEP1.py                    # Monitoring: download, filter, extract SAP numbers
├── STEP2.py                    # IW75: process report, build keys, prepare SQ01 list
├── STEP3.py                    # Contract Detail: merge data, generate final RMR file
│
├── STEP1.bat                   # Double-click launcher for Step 1
├── STEP2.bat                   # Double-click launcher for Step 2
├── STEP3.bat                   # Double-click launcher for Step 3
├── INSTALL_DEPENDENCIES.bat    # Run once to install Python packages
│
├── config.example.json         # Configuration template (copy to config.json and fill in)
│
├── input/                      # Input files go here (SAP exports + Monitoring)
└── output/                     # Generated RMR files appear here
```

---

## Setup (First Time Only)

### 1. Prerequisites

- Python 3.8+ installed with **"Add Python to PATH"** checked
- Windows OS
- Access to SAP Fiori (IW75, SQ01)
- Microsoft 365 account with access to the Monitoring SharePoint file

### 2. Install Python packages

Double-click `INSTALL_DEPENDENCIES.bat`. Wait for all `[OK]` confirmations.

### 3. Configure

```bash
# Copy the example config and fill in your values
copy config.example.json config.json
```

Open `config.json` with Notepad and update:
- `sharepoint_url` — link to the Monitoring Excel file on SharePoint
- `file_guid` — the GUID from the SharePoint URL (between `%7B` and `%7D`)
- `file_owner_path` — SharePoint personal path of the file owner
- `tenant_id` — your Microsoft 365 tenant ID
- `base_folder` — local path to this project folder

### 4. Create input/output folders

```
rmr-activation/
├── input/     ← create this folder
└── output/    ← create this folder
```

---

## Daily Usage

| Step | Action |
|------|--------|
| `STEP1.bat` | Run first. Filters Monitoring and copies SAP numbers to clipboard |
| SAP IW75 | Paste numbers → run report → export Excel to `input/` |
| `STEP2.bat` | Processes IW75, copies SQ01 list to clipboard |
| SAP SQ01 | Paste list → run report → export Excel to `input/` |
| `STEP3.bat` | Generates final `output/RMR_[date].xlsx`, opens automatically |

### File naming convention (required for auto-detection)

| File | Required format | Example |
|------|----------------|---------|
| Monitoring | `Monitoring [Month].xlsx` | `Monitoring April.xlsx` |
| IW75 export | `IW75 [Month] [Date].xlsx` | `IW75 April 4.21.2026.xlsx` |
| Contract Detail | `CONTRACT DETAIL [Date].xlsx` | `CONTRACT DETAIL 4.21.2026.xlsx` |
| Output | `RMR_[Date].xlsx` (auto-generated) | `RMR_4.21.2026.xlsx` |

---

## Monthly Update

At the start of each month, update `config.json` with the new SharePoint link for the Monitoring file. Extract the new GUID from between `%7B` and `%7D` in the URL.

---

## Key Features

- **Auto-detection** of the previous day's sheet in Monitoring
- **Automatic SharePoint download** via Microsoft 365 OAuth (falls back to manual instructions)
- **Token caching** — avoids repeated logins
- **Auto-detection of SAP header row** — handles title rows SAP adds before real data
- **Dash expansion** — `123-456` in Legacy Contract Number is split into two separate rows
- **Clipboard integration** — numbers ready to paste directly into SAP
- **Formatted Excel output** — color-coded headers, alternating rows, totals summary

---

## Configuration Reference

| Key | Description |
|-----|-------------|
| `sharepoint_url` | Full SharePoint link to the Monitoring file |
| `file_guid` | File GUID extracted from the SharePoint URL |
| `file_owner_path` | SharePoint personal site path of the file owner |
| `tenant_id` | Microsoft 365 tenant ID |
| `sharepoint_host` | SharePoint host domain |
| `base_folder` | Local base folder path |

---

## Author

**Héctor Gallego** — [@davidscript7](https://github.com/davidscript7)

---

*Built to reduce daily manual workload through Python automation integrated with SAP Fiori and Microsoft 365.*
