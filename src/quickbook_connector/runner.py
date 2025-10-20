"""
Runner file to orchestrate Excel vs QuickBooks customer comparison
and generate a JSON report.
"""

import sys
import os
from pathlib import Path
from typing import Dict, List
from datetime import datetime
import json

# ✅ Allow import from parent directory without changing structure
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# ✅ Import core logic from customer_excel_qb_sync.py (which lives outside 'src')
from customer_excel_qb_sync import (
    read_customers_from_excel,
    get_qb_customers,
    process_customers
)

DEFAULT_REPORT_NAME = "customer_sync_report.json"

# ✅ Helper: timestamp formatter
def iso_timestamp() -> str:
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")


# ✅ JSON writer
def write_report(data: Dict[str, object], file_path: Path) -> None:
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)


def run_customer_sync(
    company_file_path: str,
    excel_file_path: str,
    *,
    output_path: str | None = None,
) -> Path:
    """
    Entry point to compare Excel customers with QuickBooks and generate a JSON report.

    Args:
        company_file_path: Path to QBW company file ("" means use active).
        excel_file_path: Path to the Excel file.
        output_path: Optional custom path for JSON report.

    Returns:
        Path to the generated JSON report.
    """

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)

    report_payload: Dict[str, object] = {
        "status": "success",
        "generated_at": iso_timestamp(),
        "excel_customers_count": 0,
        "qb_customers_count": 0,
        "added_customers": [],
        "conflicts": [],
        "error": None,
    }

    try:
        print("🔄 Reading customers from Excel...")
        excel_customers = read_customers_from_excel(Path(excel_file_path))
        report_payload["excel_customers_count"] = len(excel_customers)

        print("🔌 Fetching customers from QuickBooks...")
        qb_customers = get_qb_customers()
        report_payload["qb_customers_count"] = len(qb_customers)

        print("⚙️ Running comparison and sync process...")
        result = process_customers(str(excel_file_path))

        # result should contain: added_customers, conflicts (adjust based on your actual return model)
        report_payload["added_customers"] = [
    {"name": c.name, "term": c.term, "customer_id": c.customer_id}
    for c in result.only_in_excel
]

        report_payload["conflicts"] = [
    {
        "excel_name": excel_name,
        "qb_name": qb_name,
        "term": term,
        "customer_id": cid
    }
    for (excel_name, qb_name, term, cid) in result.same_id_diff_data
]

        report_payload["only_in_qb"] = [
    {"name": c.name, "term": c.term, "customer_id": c.customer_id}
    for c in result.only_in_qb
]
        

    except Exception as exc:
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    print(f"📁 Writing JSON report to: {report_path}")
    write_report(report_payload, report_path)

    print("✅ Done.")
    return report_path


if __name__ == "__main__":
    # Example usage (you can pass real values or hook this to CLI args)
    run_customer_sync(
        company_file_path="",  # leave empty to use currently open QuickBooks company file
        excel_file_path="C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"
    )

