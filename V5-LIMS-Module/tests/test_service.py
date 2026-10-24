"""Invoicing one TR end to end, in the host's transaction (D-83)."""

from datetime import date
from decimal import Decimal

import pytest
from sqlalchemy.exc import IntegrityError

from slim_invoicing import InvoiceRefused, invoice, sales_order_for, sales_orders, save_quote
from slim_invoicing.persistence.repositories import default_quotes, quotes_for_customer, save_sales_order
from tests.factories import ACME, ANIONS, ELEMENTS, FIVE, ND, PH, TODAY, line, quote, request, sample
from tests.test_pricing import CUSTOMER, DEFAULTS


@pytest.fixture
def quoted(session):
    save_quote(session, CUSTOMER)
    save_quote(session, DEFAULTS)
    return session


def test_quotes_round_trip(quoted):
    [mine] = quotes_for_customer(quoted, ACME)
    assert mine.customer_ids == (ACME,) and mine.payment_terms == 45
    assert mine.line_items[0].analysis_ids == (ELEMENTS, ANIONS) and mine.line_items[0].bulk_price_min == 10
    assert mine.line_items[0].price_for(FIVE).bulk_price is None
    assert [len(q.line_items) for q in default_quotes(quoted)] == [3]
    assert quotes_for_customer(quoted, 999) == []


def test_a_complete_report_becomes_one_sales_order(quoted):
    order = invoice(quoted, request(sample(1, ELEMENTS, ANIONS), sample(2, PH)), "SUE", on=TODAY)
    assert order.id and order.tr_number == "TR00042" and order.payment_terms == 45 and order.created_at
    assert [(li.item, li.quantity, li.rate) for li in order.line_items] == [
        ("36 Elements + 7 Anions - Next Day", 1, Decimal("85.00")),
        ("pH - Next Day", 1, Decimal("18.00")),
    ]
    assert order.total == Decimal("103.00") and order.line_items[0].samples == ("W-001",)
    assert sales_order_for(quoted, "TR00042") == order and sales_orders(quoted) == [order]


def test_defaults_alone_use_the_labs_payment_terms(quoted):
    assert invoice(quoted, request(sample(1, PH)), "SUE", on=TODAY).payment_terms == 30


@pytest.mark.parametrize(
    ("req", "message"),
    [
        (request(sample(1, PH), complete=False), "only a complete report is invoiced"),
        (request(sample(1, PH), reported=set()), "nothing ordered was reported"),
        (request(sample(1, PH, pt=FIVE)), r"no valid quote prices analyses \[3\]"),
    ],
)
def test_what_cant_be_invoiced_is_refused(quoted, req, message):
    with pytest.raises(InvoiceRefused, match=message):
        invoice(quoted, req, "SUE", on=TODAY)
    assert sales_orders(quoted) == []


def test_a_tr_is_never_invoiced_twice(quoted):
    first = invoice(quoted, request(sample(1, PH)), "SUE", on=TODAY)
    with pytest.raises(InvoiceRefused, match="already invoiced"):
        invoice(quoted, request(sample(1, PH)), "SUE", on=TODAY)
    with pytest.raises(IntegrityError):  # and the database says so too, whoever asks
        save_sales_order(quoted, first)


def test_a_quote_needs_customers_unless_it_is_a_default():
    with pytest.raises(ValueError, match="needs at least one customer"):
        quote(line(PH, prices=[(ND, "1", None)]))
    with pytest.raises(ValueError, match="expire before"):
        quote(line(PH, prices=[(ND, "1", None)]), default=True, start=date(2027, 1, 1), end=date(2026, 1, 1))
