"""Tests for Customer Excel processing functions.

This module tests the core customer functionality and Excel-QuickBooks comparison.
"""

import tempfile
from pathlib import Path
from unittest.mock import Mock, patch
import pytest
from openpyxl import Workbook

from quickbook_connector.customer_excel_qb_sync import (
    Customer,
    CustomerComparison,
    compare_customers,
    read_customers_from_excel,
    save_customers_to_quickbooks,
    process_customers,
    create_customers_batch_qbxml,
)


def create_customers_excel(file_path: str) -> None:
    """Create a test Excel file with customer data."""
    workbook = Workbook()
    workbook.remove(workbook.active)
    sheet = workbook.create_sheet("customers")

    sheet["A1"] = "Name"
    sheet["B1"] = "Term"
    sheet["C1"] = "ID"

    customer_data = [
        ("ABC-Munster Indiana, INC. - Plant 1", "Net 45", 1),
        ("Air-O'Fallon", "None", 2),
        ("XYZ-Chicago", "Net 30", 3),
    ]

    for i, (name, term, cid) in enumerate(customer_data, start=2):
        sheet[f"A{i}"] = name
        sheet[f"B{i}"] = term
        sheet[f"C{i}"] = cid

    workbook.save(file_path)


class TestCustomers:
    """Test cases for customer functionality."""

    @pytest.fixture
    def customers_excel_file(self):
        """Create a temporary Excel file with customers for testing."""
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_path = Path(tmp.name)
        try:
            tmp.close()
            create_customers_excel(str(tmp_path))
            yield str(tmp_path)
        finally:
            try:
                if tmp_path.exists():
                    tmp_path.unlink()
            except PermissionError:
                pass

    def test_customer_dataclass(self):
        """Test Customer dataclass."""
        cust = Customer(name="ABC", term="Net 30", customer_id=10)
        assert cust.name == "ABC"
        assert cust.term == "Net 30"
        assert cust.customer_id == 10

    def test_read_customers_from_excel(self, customers_excel_file):
        """Test reading customers from Excel file."""
        customers = read_customers_from_excel(customers_excel_file)
        assert len(customers) == 3
        assert customers[0].name == "ABC-Munster Indiana, INC. - Plant 1"
        assert customers[0].term == "Net 45"
        assert customers[0].customer_id == 1
        assert customers[1].name == "Air-O'Fallon"
        assert customers[1].term == "None"
        assert customers[1].customer_id == 2

    def test_create_customers_batch_qbxml(self):
        """Test QBXML batch generation for customers."""
        customers = [
            Customer(name="ABC", term="Net 30", customer_id=1),
            Customer(name="XYZ", term="Net 45", customer_id=2),
        ]
        qbxml = create_customers_batch_qbxml(customers)

        assert "<?xml version=" in qbxml
        assert "<CustomerAdd>" in qbxml
        assert "<Name>ABC</Name>" in qbxml
        assert "<Fax>1</Fax>" in qbxml
        assert "<Name>XYZ</Name>" in qbxml
        assert "<Fax>2</Fax>" in qbxml

    @patch("quickbook_connector.customer_excel_qb_sync.win32com.client.Dispatch")
    def test_save_customers_to_quickbooks_success(self, mock_dispatch):
        """Test successful save to QuickBooks."""
        mock_qb_app = Mock()
        mock_session = "test_session"
        mock_qb_app.BeginSession.return_value = mock_session
        mock_qb_app.ProcessRequest.return_value = (
            '<?xml version="1.0"?><QBXML><QBXMLMsgsRs>'
            '<CustomerAddRs statusCode="0"><Name>ABC</Name></CustomerAddRs>'
            '<CustomerAddRs statusCode="0"><Name>XYZ</Name></CustomerAddRs>'
            '</QBXMLMsgsRs></QBXML>'
        )
        mock_dispatch.return_value = mock_qb_app

        customers = [
            Customer(name="ABC", term="Net 30", customer_id=1),
            Customer(name="XYZ", term="Net 45", customer_id=2),
        ]

        result = save_customers_to_quickbooks(customers)

        assert len(result) == 2
        assert "ABC" in result
        assert "XYZ" in result
        mock_qb_app.OpenConnection.assert_called_once()
        mock_qb_app.BeginSession.assert_called_once()
        mock_qb_app.EndSession.assert_called_once()
        mock_qb_app.CloseConnection.assert_called_once()

    @patch("quickbook_connector.customer_excel_qb_sync.win32com.client.Dispatch")
    def test_save_customers_to_quickbooks_connection_error(self, mock_dispatch):
        """Test handling QuickBooks connection error."""
        mock_dispatch.side_effect = Exception("QuickBooks not running")
        customers = [Customer(name="ABC", term="Net 30", customer_id=1)]

        with pytest.raises(RuntimeError, match="Failed to connect to QuickBooks"):
            save_customers_to_quickbooks(customers)

    @patch("quickbook_connector.customer_excel_qb_sync.get_qb_customers")
    @patch("quickbook_connector.customer_excel_qb_sync.save_customers_to_quickbooks")
    def test_process_customers_workflow(self, mock_save, mock_get_qb, customers_excel_file):
        """Test the complete customer processing workflow."""
        mock_get_qb.return_value = [
            Customer(name="ABC-Munster Indiana, INC. - Plant 1", term="Net 45", customer_id=1),
            Customer(name="Different Name", term="Net 30", customer_id=2),  # Same ID, different data
        ]
        mock_save.return_value = ["Air-O'Fallon", "XYZ-Chicago"]

        result = process_customers(customers_excel_file)

        assert isinstance(result, CustomerComparison)
        assert result.matching_count == 1
        assert len(result.same_id_diff_data) == 1
        assert len(result.only_in_excel) == 1
        assert result.only_in_excel[0].name == "XYZ-Chicago"
        mock_save.assert_called_once()

    def test_compare_customers_logic(self):
        """Test customer comparison logic separately."""
        excel_customers = [
            Customer(name="ABC", term="Net 30", customer_id=1),
            Customer(name="XYZ", term="Net 45", customer_id=2),
            Customer(name="New", term="Net 60", customer_id=3),
        ]
        qb_customers = [
            Customer(name="ABC", term="Net 30", customer_id=1),  # Same
            Customer(name="Diff", term="Net 45", customer_id=2),  # Same ID, different data
            Customer(name="Only QB", term="Net 15", customer_id=99),  # Only in QB
        ]

        result = compare_customers(excel_customers, qb_customers)

        assert result.matching_count == 1
        assert len(result.same_id_diff_data) == 1
        assert result.same_id_diff_data[0][0] == "XYZ"
        assert len(result.only_in_excel) == 1
        assert result.only_in_excel[0].name == "New"
        assert len(result.only_in_qb) == 1
        assert result.only_in_qb[0].name == "Only QB"
