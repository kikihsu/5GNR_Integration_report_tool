# config.py
# Centralized configuration for KPI Log File Processing Tool
# Edit this file to change KPI thresholds, sheet names, or file paths.

# ============================================================================
# FILE / PATH SETTINGS
# ============================================================================

LOG_FOLDER = "Logs"
OUTPUT_EXCEL = "KPI_template.xlsx"

# ============================================================================
# SHEET NAME SETTINGS
# ============================================================================

SHEET_NAMES = {
    "kept":    "Sheet1",   # Matched / kept data
    "filter":  "Sheet2",   # Cell ID filter list (user-maintained, read-only)
    "deleted": "Sheet3",   # Non-matched data
    "form":    "Sheet4",   # Generated summary form
}

# ============================================================================
# KPI RULES
# Each entry: (excel_column, display_name, operator, threshold)
#   operator must be one of: "<", ">", "<=", ">="
#
# Pass criteria:
#   PUSCH_RSSI  < -110   (lower is better, less interference)
#   ENDC_Succ   > 95%
#   RRC_CU      > 0
#   HO_Inter    > 95%
# ============================================================================

KPI_RULES = [
    ("5G_1_PUSCH_RSSI_110",  "干擾PUSCH_RSSI",            "<",  -110),
    ("5G_2_ENDC_Succ_95",    "5G ENDC建立成功率",           ">",    95),
    ("5G_3_RRC_CU",          "5G RRU CU",                  ">",     0),
    ("5G_4_HO_Inter_gNB_95", "5G HO inter gNB successful", ">",    95),
]

# ============================================================================
# ENGINEERING CATEGORY OPTIONS (used in form generation prompt)
# ============================================================================

ENGINEERING_CATEGORIES = {
    "1": "New Site",
    "2": "Add N21",
    "3": "Add N35",
}

# ============================================================================
# STYLE CONSTANTS
# ============================================================================

STYLE = {
    "header_fill":      "61cbf3",   # Blue — Sheet1/Sheet3 column headers
    "error_fill":       "C00000",   # Dark red — KPI fail cells
    "error_font":       "FFFFFF",   # White text on error cells
    "form_pass_fill":   "b5e6a2",   # Green — pass cells in Sheet4
    "form_header_fill": "dae9f8",   # Light blue — form column headers
    "form_crit_font":   "ff0000",   # Red — criterion text in form headers
    "form_note_font":   "0000ff",   # Blue — threshold hints in form rows
}
