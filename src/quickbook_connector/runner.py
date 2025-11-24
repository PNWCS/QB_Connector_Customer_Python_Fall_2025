from __future__ import annotations

from pathlib import Path
from typing import List

from quickbook_connector import compare, excel_reader, qb_gateway
from quickbook_connector.model import ComparisonReport, Customer, Conflict
from quickbook_connector.report import write_report_to_json, iso_timestamp

DEFAULT_REPORT_NAME = "customer_report.json"


def run_customer_sync(
    company_file_path: str,
    workbook_path: str,
    *,
    output_path: str | None = None,
) -> Path:
    """Synchronise customers and generate JSON report with mutual count, conflicts, and added customers."""

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)

    try:
        # 1. Read customers from Excel
        excel_customers = excel_reader.extract_customers(Path(workbook_path))

        # 2. Fetch existing customers from QuickBooks
        qb_customers = qb_gateway.fetch_customers(company_file_path)

        # 3. Compare datasets
        comparison: ComparisonReport = compare.compare_customers(
            excel_customers, qb_customers
        )

        # 4. Identify mutual customers (same ID, same name)
        mutual_customers = [
            c
            for c in excel_customers
            if c.record_id in {qb.record_id for qb in qb_customers}
            and c.name in {qb.name for qb in qb_customers}
        ]

        # 5. Add Excel-only customers to QuickBooks
        added_customers: List[Customer] = qb_gateway.add_customer_batch(
            company_file_path,
            comparison.excel_only,
        )

        # 6. Build simplified report payload
        report_payload = {
            "status": "success",
            "timestamp": iso_timestamp(),
            "added_customers": [
                {"record_id": c.record_id, "name": c.name, "source": "excel"}
                for c in added_customers
            ],
            "conflicts": [
                {
                    "record_id": c.record_id,
                    "excel_name": c.excel_name,
                    "qb_name": c.qb_name,
                    "reason": c.reason,
                }
                for c in comparison.conflicts
            ],
            "same_customers": len(mutual_customers),
            "error": None,
        }

        # 7. Write JSON report
        Path(report_path).parent.mkdir(parents=True, exist_ok=True)
        import json

        with report_path.open("w", encoding="utf-8") as f:
            json.dump(report_payload, f, indent=2)

    except Exception as exc:
        # Error report
        error_payload = {
            "status": "error",
            "timestamp": iso_timestamp(),
            "same_customers": 0,
            "added_customers": [],
            "conflicts": [],
            "error": str(exc),
        }
        Path(report_path).parent.mkdir(parents=True, exist_ok=True)
        import json

        with report_path.open("w", encoding="utf-8") as f:
            json.dump(error_payload, f, indent=2)

    return report_path


if __name__ == "__main__":
    COMPANY_FILE = ""  # Use currently open QuickBooks company file
    WORKBOOK_PATH = "C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"
    OUTPUT_PATH = "customer_report.json"

    result_path = run_customer_sync(
        company_file_path=COMPANY_FILE,
        workbook_path=WORKBOOK_PATH,
        output_path=OUTPUT_PATH,
    )

    print(f"Customer sync completed. Report generated at: {result_path}")
