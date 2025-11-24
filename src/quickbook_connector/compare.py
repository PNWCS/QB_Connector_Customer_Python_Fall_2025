from __future__ import annotations
from typing import Dict, Iterable

from quickbook_connector.model import Customer, Conflict, ComparisonReport


def compare_customers(
    excel_customers: Iterable[Customer],
    qb_customers: Iterable[Customer],
) -> ComparisonReport:
    """Compare Excel and QuickBooks customers and detect discrepancies."""

    excel_by_id: Dict[str, Customer] = {c.record_id: c for c in excel_customers}
    qb_by_id: Dict[str, Customer] = {c.record_id: c for c in qb_customers}

    # Customers with matching record_id in both sources
    mutual_ids = excel_by_id.keys() & qb_by_id.keys()

    # Conflicts: same ID but different names
    conflicts: list[Conflict] = []
    for rid in mutual_ids:
        excel_c = excel_by_id[rid]
        qb_c = qb_by_id[rid]
        if excel_c.name != qb_c.name:
            conflicts.append(
                Conflict(
                    record_id=rid,
                    excel_name=excel_c.name,
                    qb_name=qb_c.name,
                    reason="data_mismatch",
                )
            )
    
    for rid in qb_by_id:
        if rid not in excel_by_id:
            qb_c = qb_by_id[rid]
            conflicts.append(
                Conflict(
                    record_id=rid,
                    excel_name=None,
                    qb_name=qb_c.name,
                    reason="missing_in_excel",
                )
            )

    # Only in Excel
    excel_only = [cust for rid, cust in excel_by_id.items() if rid not in qb_by_id]

    # Only in QuickBooks
    qb_only = [cust for rid, cust in qb_by_id.items() if rid not in excel_by_id]

    return ComparisonReport(
        excel_only=excel_only,
        qb_only=qb_only,
        conflicts=conflicts,
    )


__all__ = ["compare_customers"]
