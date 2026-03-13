# KPI Log File Processing Tool

## 📋 Overview

This tool reads 5G KPI log files, filters them against a cell ID list, checks KPI thresholds, and generates a formatted Excel report — all automatically.

---

## 📦 What's in This Package

```
KPI_Tool/
├── main.py                 # Main program (do not edit)
├── function.py             # Core logic (do not edit)
├── config.py               # Settings: KPI thresholds, sheet names (edit if needed)
├── KPI_template.xlsx       # Excel template (update Sheet2 before each run)
├── INSTALL.bat             # ← Run this once to set up the environment
├── RUN.bat                 # ← Run this every time to start the tool
├── requirements.txt        # Library list used by INSTALL.bat
└── Logs/                   # Put your .log files here before running
    └── (your .log files)
```

---

## 🚀 First-Time Setup (Do This Once)

### Step 1 — Install Python

1. Download Python from: **https://www.python.org/downloads/**
2. Run the installer
3. ✅ **Important:** On the first screen, check **"Add Python to PATH"** before clicking Install

### Step 2 — Install Required Libraries

1. Double-click **`INSTALL.bat`**
2. A black window will open and show installation progress
3. Wait until you see **"Installation complete"**
4. Close the window

> If you see any red error text, take a screenshot and contact the tool owner.

---

## 📋 Before Each Run — Prepare Your Files

### 1. Update the filter list in `KPI_template.xlsx`

Open `KPI_template.xlsx` and go to **Sheet2**. This sheet controls which cell IDs the tool keeps.

| Column | What to fill in |
|---|---|
| `開台日期` | Activation date |
| `5G site ID` | Site ID |
| `Site Name` | Site name |
| `5G cell ID` | Cell ID to keep |

> ⚠️ Do not change the column headers. Paste your data starting from **row 2**.

### 2. Add your log files

Copy all `.log` files into the **`Logs`** folder. The tool will process every `.log` file it finds there.

> ⚠️ Make sure `KPI_template.xlsx` is **closed** in Excel before running.

---

## ▶️ Running the Tool

Double-click **`RUN.bat`**.

A window will open and walk you through the following steps automatically:

| Step | What happens |
|---|---|
| 1 | Checks that the `Logs` folder and template file exist |
| 2 | Reads all `.log` files and extracts KPI data |
| 3 | Filters rows against the cell IDs in Sheet2 |
| 4 | Writes matched data to **Sheet1**, unmatched to **Sheet3** |
| 5 | Highlights cells that fail KPI thresholds in red |
| 6 | Asks you to assign an engineering category to each site |

### Engineering category prompt

For each site, you'll see a prompt like:

```
[1/3] SITE_001 → 
```

Type the number for that site's category and press Enter:

| Input | Category |
|---|---|
| `1` | New Site |
| `2` | Add N21 |
| `3` | Add N35 |

The tool will not continue until a valid number is entered.

---

## ✅ Execution Summary

When all steps are complete, the tool displays a summary before asking optional questions:

```
  [✓] Successfully processed: 3/3 files
  [ℹ] Total rows collected: 147
  [✓] Kept (matched): 112 rows
  [ℹ] Deleted (non-matched): 35 rows
  [✓] KPI errors marked in red
  [✓] Form generated in Sheet4
```

Review this screen to confirm the run completed as expected — particularly the matched vs. unmatched row counts. If the matched count is unexpectedly low, check that Sheet2 in the template is up to date.

---

## 📊 Optional Steps After the Summary

### Generate a final report (recommended)

```
Generate report? (y/n):
```

Type `y` and press Enter. The tool creates a date-stamped file (e.g. `0115_KPI.xlsx`) in the same folder, combining the form and data into a submission-ready report. The file opens automatically when done.

### Reset the template for next use

```
Reset template for next use? (y/n):
```

Type `y` to clear Sheet1, Sheet3, and Sheet4 from the template, leaving it clean for the next run. This only works if a report was generated in the step above.

---

## 📊 KPI Pass/Fail Criteria

| KPI | Column in Sheet1 | Pass Condition | Fail Formatting |
|---|---|---|---|
| PUSCH RSSI | `5G_1_PUSCH_RSSI_110` | `< -110 dBm` | Red background, white text |
| ENDC Success Rate | `5G_2_ENDC_Succ_95` | `> 95%` | Red background, white text |
| RRC CU | `5G_3_RRC_CU` | `> 0` | Red background, white text |
| HO Inter gNB | `5G_4_HO_Inter_gNB_95` | `> 95%` | Red background, white text |

These thresholds are defined in `config.py` and can be adjusted by the tool owner if criteria change.

---

## 🐛 Troubleshooting

| Problem | What to check |
|---|---|
| `"Log folder not found"` | Create a `Logs` folder in the same directory as `RUN.bat` and add your `.log` files |
| `"Template file not found"` | Make sure `KPI_template.xlsx` is in the same folder as `RUN.bat` |
| `"Module not found"` | Re-run `INSTALL.bat` |
| Matched rows = 0 | Check that Sheet2 in `KPI_template.xlsx` has the correct cell IDs in the `5G cell ID` column |
| Garbled text in output | The tool auto-detects file encoding — install `chardet` for better accuracy: re-run `INSTALL.bat` |
| Excel file is open error | Close `KPI_template.xlsx` in Excel before running the tool |
| Window closes immediately | Right-click `RUN.bat` → **Run as administrator**, or open a Command Prompt and run `python main.py` to see the error |

---

## 💡 Tips

- You can put multiple `.log` files in the `Logs` folder and the tool processes them all in one run
- The original `.log` files are never modified or deleted
- To change KPI thresholds or engineering category options, ask the tool owner to update `config.py`

---

## 📜 Version History

| Version | Date | Notes |
|---|---|---|
| v1.1 | 2026-02-25 | Refactored: centralized config, removed manual review pauses, consistent KPI logic |
| v1.0 | 2025-01-15 | Initial release with multi-encoding support |
| v0.9 | 2025-08-07 | Original version |

