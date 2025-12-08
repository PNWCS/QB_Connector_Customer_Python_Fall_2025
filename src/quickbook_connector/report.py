from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from quickbook_connector.model import ComparisonReport, Customer, Conflict


def iso_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def _serialise_conflict(conflict: Conflict) -> Dict[str, Any]:
    return {
        "record_id": conflict.record_id,
        "excel_name": conflict.excel_name,
        "qb_name": conflict.qb_name,
        "reason": conflict.reason,
    }


def _serialise_customer(customer: Customer) -> Dict[str, Any]:
    return {
        "record_id": customer.record_id,
        "name": customer.name,
    }


def build_report_payload(
    comparison: ComparisonReport,
    mutual_data_count: int,
) -> Dict[str, Any]:
    """Build JSON payload including added customers."""

    return {
        "status": "success",
        "timestamp": iso_timestamp(),
        "mutual_data_count": mutual_data_count,
        "conflicts": [_serialise_conflict(c) for c in comparison.conflicts],
        "added_customers": [_serialise_customer(c) for c in comparison.excel_only],
    }


def write_report_to_json(
    comparison: ComparisonReport,
    mutual_data_count: int,
    output_path: Path,
) -> Path:
    payload = build_report_payload(comparison, mutual_data_count)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    return output_path
