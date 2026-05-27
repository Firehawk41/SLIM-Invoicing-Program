import csv
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem
from slim.domain.tr.enums import ProcessingTime
from slim.pipeline.sales_order_csv_writer import SalesOrderCsvWriter


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_line_item(**kwargs) -> SalesOrderLineItem:
    defaults = dict(
        id=1,
        item="ICP-MS - Next Day",
        description="36 trace elements - Next Day",
        quantity=5,
        rate=Decimal("50.00"),
        quote_price_id=1,
        processing_time=ProcessingTime.NEXT_DAY,
    )
    return SalesOrderLineItem(**(defaults | kwargs))


def _make_so(**kwargs) -> SalesOrder:
    defaults = dict(
        id=1,
        customer_name="Acme Corp",
        customer_id=1,
        submission_id=10,
        submission_ref="form_001.xlsx",
        date_received=date(2026, 5, 1),
        po_number="PO-9999",
        payment_type=PaymentType.PO_NUMBER,
        payment_terms=30,
        requester_name="Jane Doe",
        requester_email="jane@acme.com",
        service_date=date(2026, 5, 2),
        line_items=(_make_line_item(),),
    )
    return SalesOrder(**(defaults | kwargs))


def _read_csv(path: Path) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------------------
# Headers
# ---------------------------------------------------------------------------


def test_headers_written(tmp_path: Path):
    path = tmp_path / "out.csv"
    with SalesOrderCsvWriter(path):
        pass
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)
    assert headers == [
        "Customer",
        "Date",
        "Item",
        "Description",
        "Quantity",
        "Rate",
        "PO Number",
        "Name of Requester",
        "Email of Requester",
        "Service Date",
        "Submission Ref",
    ]


# ---------------------------------------------------------------------------
# Writing rows
# ---------------------------------------------------------------------------


def test_write_single_so_single_line_item(tmp_path: Path):
    path = tmp_path / "out.csv"
    so = _make_so()
    with SalesOrderCsvWriter(path) as w:
        w.write(so)
    rows = _read_csv(path)
    assert len(rows) == 1
    assert rows[0]["Customer"] == "Acme Corp"
    assert rows[0]["Date"] == "2026-05-01"
    assert rows[0]["Item"] == "ICP-MS - Next Day"
    assert rows[0]["Quantity"] == "5"
    assert rows[0]["PO Number"] == "PO-9999"
    assert rows[0]["Submission Ref"] == "form_001.xlsx"


def test_write_produces_one_row_per_line_item(tmp_path: Path):
    path = tmp_path / "out.csv"
    so = _make_so(
        line_items=(
            _make_line_item(item="ICP-MS - Next Day"),
            _make_line_item(item="Si - Next Day"),
        )
    )
    with SalesOrderCsvWriter(path) as w:
        w.write(so)
    rows = _read_csv(path)
    assert len(rows) == 2
    items = [r["Item"] for r in rows]
    assert items == ["ICP-MS - Next Day", "Si - Next Day"]


def test_write_multiple_sales_orders(tmp_path: Path):
    path = tmp_path / "out.csv"
    so1 = _make_so(customer_name="Acme", submission_ref="a.xlsx")
    so2 = _make_so(customer_name="Beta", submission_ref="b.xlsx")
    with SalesOrderCsvWriter(path) as w:
        w.write(so1)
        w.write(so2)
    rows = _read_csv(path)
    assert len(rows) == 2
    assert rows[0]["Customer"] == "Acme"
    assert rows[1]["Customer"] == "Beta"


def test_write_no_line_items_produces_no_rows(tmp_path: Path):
    path = tmp_path / "out.csv"
    so = _make_so(line_items=())
    with SalesOrderCsvWriter(path) as w:
        w.write(so)
    rows = _read_csv(path)
    assert rows == []


# ---------------------------------------------------------------------------
# Context manager
# ---------------------------------------------------------------------------


def test_file_closed_after_context_exit(tmp_path: Path):
    path = tmp_path / "out.csv"
    writer = SalesOrderCsvWriter(path)
    with writer:
        pass
    assert writer._file is None


def test_write_outside_context_raises(tmp_path: Path):
    path = tmp_path / "out.csv"
    writer = SalesOrderCsvWriter(path)
    with pytest.raises(AssertionError, match="context manager"):
        writer.write(_make_so())
