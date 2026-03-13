# main.py
# Updated: 2026-02-25
# Main orchestration script
# Before running: update Sheet2 in KPI_template.xlsx with your cell ID list.
# NOTICE: DO NOT adjust the title row — paste data starting from row 2.
"""
KPI Log File Processing Tool
============================
Main execution script — run this file to process your log files.

Author: Kiki Hsu
"""

import os
import sys

import function
from config import (
    LOG_FOLDER, OUTPUT_EXCEL, SHEET_NAMES,
    ENGINEERING_CATEGORIES,
)


# ============================================================================
# DISPLAY HELPERS
# ============================================================================

def print_banner(text):
    print("\n" + "#" * 70)
    print("#" + " " * 68 + "#")
    print("#" + text.center(68) + "#")
    print("#" + " " * 68 + "#")
    print("#" * 70 + "\n")

def print_step(step_num, total_steps, description):
    print("\n" + "=" * 70)
    print(f"STEP {step_num}/{total_steps}: {description}")
    print("=" * 70)

def print_status(message, status="INFO"):
    symbols = {"INFO": "i", "SUCCESS": "+", "ERROR": "X", "WARNING": "!"}
    print(f"  [{symbols.get(status, '•')}] {message}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    sheet_kept    = SHEET_NAMES["kept"]
    sheet_filter  = SHEET_NAMES["filter"]
    sheet_deleted = SHEET_NAMES["deleted"]
    sheet_form    = SHEET_NAMES["form"]
    total_steps   = 6

    print_banner("KPI LOG FILE PROCESSING TOOL")

    # ── STEP 1: Validate environment ─────────────────────────────────────────
    print_step(1, total_steps, "Validating Environment")

    if not os.path.exists(LOG_FOLDER):
        print_status(f"Log folder '{LOG_FOLDER}' not found!", "ERROR")
        print_status("Please create a 'Logs' folder and add your .log files", "WARNING")
        input("\nPress Enter to exit...")
        sys.exit(1)

    log_files = [f for f in os.listdir(LOG_FOLDER) if f.endswith(".log")]
    if not log_files:
        print_status(f"No .log files found in '{LOG_FOLDER}'", "ERROR")
        input("\nPress Enter to exit...")
        sys.exit(1)

    print_status(f"Found {len(log_files)} log file(s)", "SUCCESS")
    for i, f in enumerate(log_files, 1):
        print(f"    {i}. {f}")

    if not os.path.exists(OUTPUT_EXCEL):
        print_status(f"Template '{OUTPUT_EXCEL}' not found!", "ERROR")
        print_status("Ensure KPI_template.xlsx is in the same folder as main.py", "WARNING")
        input("\nPress Enter to exit...")
        sys.exit(1)

    print_status(f"Template '{OUTPUT_EXCEL}' found", "SUCCESS")

    # ── STEP 2: Read log files ────────────────────────────────────────────────
    print_step(2, total_steps, "Reading Log Files")

    all_data     = []
    success_count = 0
    failed_files  = []

    for i, filename in enumerate(log_files, 1):
        print(f"\n  [{i}/{len(log_files)}] Processing: {filename}")
        file_path = os.path.join(LOG_FOLDER, filename)
        try:
            rows = function.read_file(file_path, filename)
            if rows:
                all_data.extend(rows)
                success_count += 1
                print_status(f"Successfully read {len(rows)} rows", "SUCCESS")
            else:
                print_status("No data extracted from file", "WARNING")
                failed_files.append(filename)
        except Exception as exc:
            print_status(f"Failed to process: {exc}", "ERROR")
            failed_files.append(filename)

    print(f"\n  Summary:")
    print_status(f"Successfully processed: {success_count}/{len(log_files)} files", "SUCCESS")
    print_status(f"Total rows collected: {len(all_data)}", "INFO")
    if failed_files:
        print_status(f"Failed files: {', '.join(failed_files)}", "WARNING")

    if not all_data:
        print_status("No data to process. Exiting.", "ERROR")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # ── STEP 3: Filter data ───────────────────────────────────────────────────
    print_step(3, total_steps, "Filtering Data")

    kept_data, deleted_data = function.filter_data(all_data, OUTPUT_EXCEL, sheet_filter)
    print_status(f"Kept (matched): {len(kept_data)} rows", "SUCCESS")
    print_status(f"Deleted (non-matched): {len(deleted_data)} rows", "INFO")

    # ── STEP 4: Export to Excel ───────────────────────────────────────────────
    print_step(4, total_steps, "Exporting to Excel")

    print_status(f"Writing kept data to '{sheet_kept}'...", "INFO")
    function.output_file(kept_data, OUTPUT_EXCEL, sheet_kept)

    print_status(f"Writing deleted data to '{sheet_deleted}'...", "INFO")
    function.output_file(deleted_data, OUTPUT_EXCEL, sheet_deleted)

    print_status("Excel file updated", "SUCCESS")

    # ── STEP 5: Mark KPI errors ───────────────────────────────────────────────
    print_step(5, total_steps, "Marking KPI Errors")

    function.check_and_format_kpi_data(OUTPUT_EXCEL, sheet_kept)
    print_status("KPI errors marked in red", "SUCCESS")

    # ── STEP 6: Collect engineering categories, then generate form ────────────
    print_step(6, total_steps, "Collecting Engineering Categories & Generating Form")

    import pandas as pd
    df_sheet2 = pd.read_excel(OUTPUT_EXCEL, sheet_name=sheet_filter, engine="openpyxl")
    unique_site_ids = df_sheet2["5G site ID"].unique()

    options_text = "  " + "  ".join(
        f"{k}.{v}" for k, v in ENGINEERING_CATEGORIES.items()
    )

    engineering_categories = {}
    print(f"\n  Please enter the engineering category for each site:")
    print(f"  Options: {options_text}\n")

    for idx, site_id in enumerate(unique_site_ids, 1):
        while True:
            choice = input(
                f"    [{idx}/{len(unique_site_ids)}] {site_id} > "
            ).strip()
            if choice in ENGINEERING_CATEGORIES:
                engineering_categories[str(site_id)] = ENGINEERING_CATEGORIES[choice]
                break
            valid = "/".join(ENGINEERING_CATEGORIES.keys())
            print(f"      Invalid input. Please enter {valid}.")

    function.process_kpi_excel(OUTPUT_EXCEL, engineering_categories)
    print_status("Form generated in Sheet4", "SUCCESS")

    # ── OPTIONAL: Generate final report ──────────────────────────────────────
    print("\n" + "=" * 70)
    print("OPTIONAL: Generate Final Report")
    print("=" * 70)

    report_choice = input("\n  Generate report? (y/n): ").strip().lower()
    file_name = None

    if report_choice == "y":
        print_status("Generating report...", "INFO")
        file_name = function.process_excel_template(OUTPUT_EXCEL)
        if file_name:
            print_status(f"Report created: {file_name}", "SUCCESS")
            function.open_file(file_name)
        else:
            print_status("Report generation failed", "ERROR")
    else:
        print_status("Skipping report generation", "INFO")

    # ── OPTIONAL: Reset template ──────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("OPTIONAL: Reset Template")
    print("=" * 70)

    reset_choice = input("\n  Reset template for next use? (y/n): ").strip().lower()

    if reset_choice == "y":
        if file_name:
            print_status("Resetting template...", "INFO")
            function.reset_excel_template(
                file_name=OUTPUT_EXCEL,
                sheet_name_kept=sheet_kept,
                sheet_name_deleted=sheet_deleted,
                sheet_name_form=sheet_form,
            )
            print_status("Template reset complete", "SUCCESS")
        else:
            print_status("Cannot reset: no report was generated", "WARNING")
    else:
        print_status("Template not reset", "INFO")

    print_banner("PROCESSING COMPLETE!")
    input("Press Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[!] Process interrupted by user")
        sys.exit(0)
    except Exception as exc:
        print(f"\n\n[X] FATAL ERROR: {exc}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)
