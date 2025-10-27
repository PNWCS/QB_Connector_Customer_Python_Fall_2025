from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest

from quickbook_connector.model import ComparisonReport, Conflict, Customer


@pytest.fixture
def mock_excel_terms():
    """Mock Excel payment terms."""
    return [
        Customer(record_id="15", name="Net 15", source="excel"),
        Customer(record_id="30", name="Net 30", source="excel"),
        Customer(record_id="45", name="Net 45", source="excel"),
    ]