"""Domain models for payment term synchronisation.

These dataclasses represent the core entities shared throughout the tool:
payment terms, conflicts between sources, and the aggregate comparison report.
"""

from __future__ import annotations  # Postponed evaluation of annotations (PEP 563)

from dataclasses import dataclass, field  # Dataclass utilities
from typing import Literal  # Constrained string types for clarity

SourceLiteral = Literal["excel", "quickbooks"]  # Origin of a PaymentTerm
ConflictReason = Literal[
    "name_mismatch", "missing_in_excel", "missing_in_quickbooks"
]  # Why a conflict exists


@dataclass(slots=True)
class Customer:
    """Represents a customer synchronised between Excel and QuickBooks."""

    record_id: str  # Unique identifier (company database id)
    name: str  # Human-readable name (e.g., "Net 30")
    source: SourceLiteral  # "excel" or "quickbooks"

    def __str__(self) -> str:
        return (
            f"customers(id={self.record_id}, name={self.name}, source={self.source})"
        )


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks payment terms."""

    record_id: str  # Shared identifier for the conflicting term
    excel_name: str | None  # Name from Excel (None if missing there)
    qb_name: str | None  # Name from QuickBooks (None if missing there)
    reason: ConflictReason  # Explanation of the conflict type


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison outcomes for later processing."""

    excel_only: list[Customer] = field(default_factory=list)  # Present only in Excel
    qb_only: list[Customer] = field(
        default_factory=list
    )  # Present only in QuickBooks
    conflicts: list[Conflict] = field(default_factory=list)  # Same ID, differing names


__all__ = [
    "customers",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]