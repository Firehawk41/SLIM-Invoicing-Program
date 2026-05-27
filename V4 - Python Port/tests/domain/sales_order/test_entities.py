from datetime import date
from decimal import Decimal

import pytest

from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem
from slim.domain.tr.enums import ProcessingTime


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_line_item(**kwargs) -> SalesOrderLineItem:
    defaults = dict(
        id=0,
        item="ICP-MS - Next Day",
        description="36 trace elements - Next Day",
        quantity=5,
        rate=Decimal("50.00"),
        quote_price_id=1,
        processing_time=ProcessingTime.NEXT_DAY,
    )
    return SalesOrderLineItem(**(defaults | kwargs))


def _make_sales_order(**kwargs) -> SalesOrder:
    defaults = dict(
        id=0,
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
    )
    return SalesOrder(**(defaults | kwargs))


# ---------------------------------------------------------------------------
# SalesOrderLineItem — validate
# ---------------------------------------------------------------------------


def test_line_item_validate_happy_path():
    assert _make_line_item().validate() is True


def test_line_item_validate_zero_quantity():
    assert _make_line_item(quantity=0).validate() is False


def test_line_item_validate_negative_quantity():
    assert _make_line_item(quantity=-1).validate() is False


def test_line_item_validate_negative_rate():
    assert _make_line_item(rate=Decimal("-0.01")).validate() is False


def test_line_item_validate_zero_rate_is_valid():
    # Zero rate is allowed (e.g. comped analysis)
    assert _make_line_item(rate=Decimal("0")).validate() is True


def test_line_item_validate_blank_item():
    assert _make_line_item(item="").validate() is False


# ---------------------------------------------------------------------------
# SalesOrderLineItem — total
# ---------------------------------------------------------------------------


def test_line_item_total():
    li = _make_line_item(quantity=3, rate=Decimal("25.50"))
    assert li.total == Decimal("76.50")


def test_line_item_total_zero_rate():
    li = _make_line_item(quantity=10, rate=Decimal("0"))
    assert li.total == Decimal("0")


# ---------------------------------------------------------------------------
# SalesOrder — validate
# ---------------------------------------------------------------------------


def test_sales_order_validate_happy_path():
    assert _make_sales_order().validate() is True


def test_sales_order_validate_zero_customer_id():
    assert _make_sales_order(customer_id=0).validate() is False


def test_sales_order_validate_blank_customer_name():
    assert _make_sales_order(customer_name="").validate() is False


def test_sales_order_validate_min_service_date():
    assert _make_sales_order(service_date=date.min).validate() is False


def test_sales_order_validate_min_date_received():
    assert _make_sales_order(date_received=date.min).validate() is False


def test_sales_order_validate_po_number_payment_type_requires_po():
    so = _make_sales_order(payment_type=PaymentType.PO_NUMBER, po_number="")
    assert so.validate() is False


def test_sales_order_validate_credit_card_no_po_is_valid():
    so = _make_sales_order(payment_type=PaymentType.CREDIT_CARD, po_number="")
    assert so.validate() is True


def test_sales_order_validate_negative_payment_terms():
    assert _make_sales_order(payment_terms=-1).validate() is False


def test_sales_order_validate_zero_payment_terms_is_valid():
    assert _make_sales_order(payment_terms=0).validate() is True


# ---------------------------------------------------------------------------
# SalesOrder — total_value
# ---------------------------------------------------------------------------


def test_sales_order_total_value_no_line_items():
    so = _make_sales_order()
    assert so.total_value == Decimal("0")


def test_sales_order_total_value_sums_line_items():
    li1 = _make_line_item(quantity=2, rate=Decimal("100.00"))
    li2 = _make_line_item(quantity=3, rate=Decimal("50.00"))
    so = _make_sales_order(line_items=(li1, li2))
    assert so.total_value == Decimal("350.00")
