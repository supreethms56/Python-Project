"""
app.py: Entry point orchestrating directory scanning and final summary metrics.
Fulfills:
- AC-1: Scan user-specified folder and identify valid .xlsx files.
- AC-2: Process multiple Excel files in a single execution.
- AC-15: Display total files processed, total revenue, and top-selling product.
"""

import sys
from pathlib import Path
from collections import defaultdict
from auditor import ReportAuditor


def audit_sales_directory(target_folder: str):
    folder_path = Path(target_folder)

    # Validate target directory
    if not folder_path.exists() or not folder_path.is_dir():
        print(f"Error: Directory '{target_folder}' does not exist or is not a folder.")
        return

    # AC-1: Automatically find all .xlsx files excluding temporary or already-audited files
    candidate_files = [
        f for f in folder_path.glob("*.xlsx")
        if not f.name.startswith("~$") and not f.name.startswith("audit_")
    ]

    if not candidate_files:
        print(f"No valid .xlsx files found in '{target_folder}'.")
        return

    print(f"Found {len(candidate_files)} file(s) to process in '{target_folder}'. Starting audit...\n")

    total_files_processed = 0
    total_revenue = 0.0
    overall_product_units = defaultdict(float)

    # AC-2: Process multiple files in a single execution
    for file_path in candidate_files:
        print(f"Auditing file: {file_path.name}")
        auditor = ReportAuditor(file_path)
        success = auditor.process_file()

        if success:
            total_files_processed += 1
            total_revenue += auditor.file_revenue
            for product, units in auditor.product_units_sold.items():
                overall_product_units[product] += units

    # Determine top-selling product by cumulative quantity
    top_selling_product = "None"
    if overall_product_units:
        top_selling_product = max(overall_product_units, key=overall_product_units.get)

    # AC-15: Print summary report to the terminal
    print("\n" + "=" * 50)
    print("              AUDIT EXECUTION SUMMARY             ")
    print("=" * 50)
    print(f" Total Files Processed : {total_files_processed}")
    print(f" Total Revenue         : ${total_revenue:,.2f}")
    print(f" Top-Selling Product   : {top_selling_product}")
    print("=" * 50 + "\n")


if __name__ == "__main__":
    # Allows folder name to be passed as a CLI argument: python app.py ./reports
    target_dir = sys.argv[1] if len(sys.argv) > 1 else "C:/Users/Supreeth/OneDrive - Aritha Consulting Services/Desktop/Dataset"
    audit_sales_directory(target_dir)