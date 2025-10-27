from __future__ import annotations

from pathlib import Path
from unittest.mock import patch, MagicMock, Mock
import pytest

from quickbook_connector.model import Customer
from quickbook_connector.excel_reader import extract_customers


@pytest.fixture
def mock_excel_terms():
    """Mock Excel payment terms."""
    return [
        Customer(record_id="15", name="Net 15", source="excel"),
        Customer(record_id="30", name="Net 30", source="excel"),
        Customer(record_id="45", name="Net 45", source="excel"),
    ]


def test_extract_customers_file_not_found(tmp_path):
    """Ensure FileNotFoundError is raised when workbook is missing."""
    missing_file = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        extract_customers(missing_file)


@patch("quickbook_connector.excel_reader.load_workbook")
def test_extract_customers_missing_worksheet(mock_load):
    """Ensure ValueError is raised when 'customers' worksheet is missing."""
    workbook_mock = MagicMock()
    mock_load.return_value = workbook_mock
    workbook_mock.__getitem__.side_effect = KeyError("Sheet not found")

    with pytest.raises(ValueError, match="Worksheet 'customers' not found"):
        extract_customers(Path("C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"))

    workbook_mock.close.assert_called_once()


@patch("quickbook_connector.excel_reader.load_workbook")
def test_extract_customers_empty_sheet(mock_load):
    """Return empty list if the worksheet has no data rows."""
    mock_sheet = MagicMock()
    mock_sheet.iter_rows.return_value = iter([])

    workbook_mock = MagicMock()
    workbook_mock.__getitem__.return_value = mock_sheet
    mock_load.return_value = workbook_mock

    result = extract_customers(Path("C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"))

    assert result == []
    workbook_mock.close.assert_called_once()


@patch("quickbook_connector.excel_reader.load_workbook")
def test_extract_customers_valid_data(mock_load):
    """Successfully extract valid Customer records from Excel."""
    # Simulated worksheet rows
    header = ("ID", "Name")
    rows = [
        (15, "Net 15"),
        (30.0, "Net 30"),
        ("45", "Net 45"),
        (None, "Invalid No ID"),
        (90, None),
        ("", "Blank ID"),
        (60, "   "),  # blank name
    ]

    mock_sheet = MagicMock()
    mock_sheet.iter_rows.return_value = iter([header] + rows)

    workbook_mock = MagicMock()
    workbook_mock.__getitem__.return_value = mock_sheet
    mock_load.return_value = workbook_mock

    result = extract_customers(Path("C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"))

    # Only valid entries should be returned
    assert len(result) == 3
    assert all(isinstance(c, Customer) for c in result)
    assert [c.record_id for c in result] == ["15", "30", "45"]
    assert [c.name for c in result] == ["Net 15", "Net 30", "Net 45"]
    assert all(c.source == "excel" for c in result)

    workbook_mock.close.assert_called_once()


@patch("quickbook_connector.excel_reader.load_workbook")
def test_extract_customers_non_integer_ids(mock_load):
    """Handle non-numeric IDs gracefully by converting to strings."""
    header = ("ID", "Name")
    rows = [("A15", "Custom Term"), ("15days", "Term 15 Days")]

    mock_sheet = MagicMock()
    mock_sheet.iter_rows.return_value = iter([header] + rows)

    workbook_mock = MagicMock()
    workbook_mock.__getitem__.return_value = mock_sheet
    mock_load.return_value = workbook_mock

    result = extract_customers(Path("C:/Users/BoyaA/Desktop/QB_Connector_Customer_Python_Fall_2025/company_data.xlsx"))

    assert len(result) == 2
    assert result[0].record_id == "A15"
    assert result[1].record_id == "15days"
    assert result[0].source == "excel"
    workbook_mock.close.assert_called_once()
