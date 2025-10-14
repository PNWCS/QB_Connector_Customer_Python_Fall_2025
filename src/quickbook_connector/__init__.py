"""Customer Excel-QB Sync Package.

A module for reading customer data from Excel, comparing with QuickBooks,
and synchronizing differences automatically.
"""

from .customer_excel_qb_sync import (
    Customer,
    CustomerComparison,
    read_customers_from_excel,
    get_qb_customers,
    compare_customers,
    process_customers,
)

__version__ = "0.1.0"
__all__ = [
    "Customer",
    "CustomerComparison",
    "read_customers_from_excel",
    "get_qb_customers",
    "compare_customers",
    "process_customers",
]
