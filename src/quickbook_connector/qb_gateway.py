"""QuickBooks COM gateway helpers for customer terms.

This module communicates with QuickBooks Desktop via the QBXML Request Processor
COM interface exposed by ``pywin32``. It provides functions to query existing
customer terms and to add new terms either individually or in batches.
"""

from __future__ import annotations  # Enable postponed evaluation of annotations

import xml.etree.ElementTree as ET  # XML parsing for QBXML responses
from contextlib import contextmanager  # For clean session management
from typing import Iterator, List  # Type hints for readability

try:
    import win32com.client  # type: ignore  # Imported lazily to allow testing without pywin32
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore  # Fallback used to raise a clear error later

from quickbook_connector.model import Customer  # Core domain model for customer terms


APP_NAME = "Quickbooks Connector"  # do not chanege this (intentional spelling per spec)


def _require_win32com() -> None:
    """Ensure the win32com dependency is available before COM operations."""
    if win32com is None:  # pragma: no cover - exercised via tests
        raise RuntimeError("pywin32 is required to communicate with QuickBooks")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    """Context manager that opens and closes a QBXML request processor session.

    Yields a tuple of (session, ticket) that must be used to send requests.
    Ensures sessions are properly closed even if errors occur.
    """
    _require_win32com()
    session = win32com.client.Dispatch(
        "QBXMLRP2.RequestProcessor"
    )  # Acquire COM object
    session.OpenConnection2("", APP_NAME, 1)  # Register connection with an app name
    ticket = session.BeginSession(
        "", 0
    )  # Start a session using the currently open company file
    try:
        yield session, ticket  # Provide the session and ticket to the caller
    finally:
        try:
            session.EndSession(ticket)  # Always end the session
        finally:
            session.CloseConnection()  # And close the connection regardless of errors


def _send_qbxml(qbxml: str) -> ET.Element:
    """Send a QBXML request and return the parsed XML root element."""
    with _qb_session() as (session, ticket):  # Debug output to aid diagnostics
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
    return _parse_response(raw_response)


def _parse_response(raw_xml: str) -> ET.Element:
    """Parse raw QBXML response and raise on error status codes."""
    root = ET.fromstring(raw_xml)  # Parse the XML text into an element tree
    response = root.find(".//*[@statusCode]")  # Locate the first node with a status
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))  # Convert status code to int
    status_message = response.get("statusMessage", "")  # Retrieve status message
    # Status code 1 means "no matching objects found" - this is OK for queries
    if status_code != 0 and status_code != 1:
        print(f"QuickBooks error ({status_code}): {status_message}")  # Log the error
        raise RuntimeError(status_message)  # Propagate as a runtime error
    return root  # Return the parsed XML root


def fetch_customers(company_file: str | None = None) -> List[Customer]:
    """Return customers currently stored in QuickBooks.

    The ``company_file`` parameter is currently unused since sessions default to
    the open company file, but is kept for API symmetry with other functions.
    """

    qbxml = """\
<?xml version="1.0"?>
<?qbxml version="16.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <CustomerQueryRq/>
  </QBXMLMsgsRq>
</QBXML>"""
    root = _send_qbxml(qbxml)  # Dispatch request and parse response
    terms: List[Customer] = []  # Accumulate parsed terms here
    for term_ret in root.findall(".//CustomerRet"):  # Iterate over each term
        record_id = term_ret.findtext("Fax")  # Use days as record_id
        name = (term_ret.findtext("FullName") or "").strip()  # Normalise the name

        if not record_id:
            continue  # Skip entries without an ID
        try:
            record_id = str(
                int(record_id)
            )  # Coerce to int then back to str to normalise
        except ValueError:
            record_id = record_id.strip()  # Fallback: trim whitespace
        if not record_id:
            continue  # Skip empty IDs after normalisation

        terms.append(Customer(record_id=record_id, name=name, source="quickbooks"))

    return terms  # Return all collected terms


def add_customer_batch(
    company_file: str | None, terms: List[Customer]
) -> List[Customer]:
    """Create multiple customer terms in QuickBooks in a single batch request."""

    if not terms:
        return []  # Nothing to add; return early

    # Build the QBXML with multiple StandardTermsAddRq entries
    requests = []  # Collect individual add requests to embed in one batch
    for term in terms:
        try:
            days_value = int(term.record_id)  # QuickBooks expects a numeric days value
        except ValueError as exc:
            raise ValueError(
                f"record_id must be numeric for QuickBooks customer terms: {term.record_id}"
            ) from exc

        # Build the QBXML snippet for this term creation
        requests.append(
            f"    <StandardTermsAddRq>\n"
            f"      <StandardTermsAdd>\n"
            f"        <Name>{_escape_xml(term.name)}</Name>\n"
            f"        <StdDiscountDays>{days_value}</StdDiscountDays>\n"
            f"        <DiscountPct>0</DiscountPct>\n"
            f"      </StandardTermsAdd>\n"
            f"    </StandardTermsAddRq>"
        )

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n' + "\n".join(requests) + "\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )  # Batch request enabling partial success on errors

    try:
        root = _send_qbxml(qbxml)  # Submit the batch to QuickBooks
    except RuntimeError as exc:
        # If the entire batch fails, return empty list
        print(f"Batch add failed: {exc}")
        return []

    # Parse all responses
    added_terms: List[Customer] = []  # Terms confirmed/returned by QuickBooks
    for term_ret in root.findall(".//StandardTermsRet"):
        record_id = term_ret.findtext("StdDiscountDays")  # Extract the ID
        if not record_id:
            continue
        try:
            record_id = str(int(record_id))  # Normalise numeric string
        except ValueError:
            record_id = record_id.strip()
        name = (term_ret.findtext("Name") or "").strip()  # Extract and trim name
        added_terms.append(
            Customer(record_id=record_id, name=name, source="quickbooks")
        )

    return added_terms  # Return all terms that were added/acknowledged


def add_customer(company_file: str | None, term: Customer) -> Customer:
    """Create a single customer term in QuickBooks and return the stored record."""

    try:
        days_value = int(term.record_id)  # Validate that the ID is numeric
    except ValueError as exc:
        raise ValueError(
            "record_id must be numeric for QuickBooks customer terms"
        ) from exc

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <StandardTermsAddRq>\n"
        "      <StandardTermsAdd>\n"
        f"        <Name>{_escape_xml(term.name)}</Name>\n"
        f"        <StdDiscountDays>{days_value}</StdDiscountDays>\n"
        "        <DiscountPct>0</DiscountPct>\n"
        "      </StandardTermsAdd>\n"
        "    </StandardTermsAddRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )  # Single add request with stop-on-error behavior

    try:
        root = _send_qbxml(qbxml)  # Send request and parse response
    except RuntimeError as exc:
        # Check if error is "name already in use" (error code 3100)
        if "already in use" in str(exc):
            # Return the term as-is since it already exists
            return Customer(
                record_id=term.record_id, name=term.name, source="quickbooks"
            )
        raise  # Re-raise other errors

    term_ret = root.find(".//StandardTermsRet")  # Extract the returned record, if any
    if term_ret is None:
        # Some responses may omit the created object; fall back to input values
        return Customer(record_id=term.record_id, name=term.name, source="quickbooks")

    record_id = term_ret.findtext("StdDiscountDays") or term.record_id  # Prefer QB's ID
    try:
        record_id = str(int(record_id))  # Normalise to a clean numeric string
    except ValueError:
        record_id = record_id.strip()
    name = (term_ret.findtext("Name") or term.name).strip()  # Prefer QB's name

    return Customer(record_id=record_id, name=name, source="quickbooks")


def _escape_xml(value: str) -> str:
    """Escape XML special characters for safe QBXML construction."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


__all__ = [
    "fetch_customers",
    "add_customer",
    "add_customer_batch",
]  # Public API


if __name__ == "__main__":  # manual test run
    try:
        customers = fetch_customers()  # No need to pass Excel path
        for customer in customers:
            print(customer)
    except Exception as e:
        print(f"Error: {e}")
