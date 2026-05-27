from datetime import date
from decimal import Decimal

import pytest

from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_line_item import QuoteLineItem
from slim.domain.quote.quote_price import QuotePrice
from slim.domain.tr.enums import ProcessingTime, RequestType


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_price(**kwargs) -> QuotePrice:
    defaults = dict(
        id=0,
        quote_line_item_id=0,
        regular_price=Decimal("100.00"),
        bulk_price=Decimal("-1"),
        processing_time=ProcessingTime.NEXT_DAY,
        is_valid=True,
    )
    return QuotePrice(**(defaults | kwargs))


def _make_line_item(**kwargs) -> QuoteLineItem:
    defaults = dict(
        id=0,
        quote_id=0,
        request_type=RequestType.CHEMICAL,
        bulk_price_min=0,
        chemical_id=None,
        is_valid=True,
        prices=(_make_price(),),
    )
    return QuoteLineItem(**(defaults | kwargs))


def _make_quote(**kwargs) -> Quote:
    defaults = dict(
        id=0,
        effective_date=date(2026, 1, 1),
        expire_date=date(2026, 12, 31),
        payment_type=PaymentType.PO_NUMBER,
        payment_terms=30,
        created_by="jt",
        created_date=date(2026, 1, 1),
        is_valid=True,
        is_default_quote=False,
        customer_ids=(1,),
        line_items=(_make_line_item(),),
    )
    return Quote(**(defaults | kwargs))


# ---------------------------------------------------------------------------
# PaymentType
# ---------------------------------------------------------------------------

def test_payment_type_values() -> None:
    assert PaymentType.PO_NUMBER == 1
    assert PaymentType.CREDIT_CARD == 2


# ---------------------------------------------------------------------------
# QuotePrice
# ---------------------------------------------------------------------------

def test_quote_price_fields() -> None:
    p = _make_price(id=5, regular_price=Decimal("75.50"), processing_time=ProcessingTime.SAME_DAY_RUSH)
    assert p.id == 5
    assert p.regular_price == Decimal("75.50")
    assert p.bulk_price == Decimal("-1")
    assert p.processing_time == ProcessingTime.SAME_DAY_RUSH
    assert p.is_valid is True


def test_quote_price_frozen() -> None:
    p = _make_price()
    with pytest.raises(Exception):
        p.regular_price = Decimal("0")  # type: ignore[misc]


def test_quote_price_has_bulk_pricing_false_when_sentinel() -> None:
    assert not _make_price(bulk_price=Decimal("-1")).has_bulk_pricing


def test_quote_price_has_bulk_pricing_true_when_set() -> None:
    assert _make_price(bulk_price=Decimal("80.00")).has_bulk_pricing


# ---------------------------------------------------------------------------
# QuoteLineItem
# ---------------------------------------------------------------------------

def test_quote_line_item_fields() -> None:
    li = _make_line_item(id=3, bulk_price_min=5, chemical_id=7, analysis_ids=(10, 20))
    assert li.id == 3
    assert li.bulk_price_min == 5
    assert li.chemical_id == 7
    assert li.analysis_ids == (10, 20)
    assert li.element_ids == ()
    assert len(li.prices) == 1


def test_quote_line_item_frozen() -> None:
    li = _make_line_item()
    with pytest.raises(Exception):
        li.bulk_price_min = 99  # type: ignore[misc]


def test_get_price_returns_matching_price() -> None:
    p_next = _make_price(processing_time=ProcessingTime.NEXT_DAY, regular_price=Decimal("100"))
    p_rush = _make_price(processing_time=ProcessingTime.SAME_DAY_RUSH, regular_price=Decimal("200"))
    li = _make_line_item(prices=(p_next, p_rush))
    result = li.get_price(ProcessingTime.SAME_DAY_RUSH)
    assert result is not None
    assert result.regular_price == Decimal("200")


def test_get_price_returns_none_when_not_found() -> None:
    li = _make_line_item(prices=(_make_price(processing_time=ProcessingTime.NEXT_DAY),))
    assert li.get_price(ProcessingTime.THREE_DAYS) is None


def test_quote_line_item_null_chemical_id() -> None:
    assert _make_line_item(chemical_id=None).chemical_id is None


# ---------------------------------------------------------------------------
# Quote
# ---------------------------------------------------------------------------

def test_quote_fields() -> None:
    q = _make_quote(id=1, payment_type=PaymentType.CREDIT_CARD, payment_terms=0)
    assert q.id == 1
    assert q.payment_type == PaymentType.CREDIT_CARD
    assert q.payment_terms == 0
    assert q.comments == ""
    assert q.modified_date is None
    assert q.is_default_quote is False


def test_quote_frozen() -> None:
    q = _make_quote()
    with pytest.raises(Exception):
        q.payment_terms = 99  # type: ignore[misc]


def test_quote_validate_passes_customer_quote() -> None:
    assert _make_quote().validate() is True


def test_quote_validate_passes_default_quote() -> None:
    q = _make_quote(is_default_quote=True, customer_ids=())
    assert q.validate() is True


def test_quote_validate_fails_no_customers_for_non_default() -> None:
    q = _make_quote(is_default_quote=False, customer_ids=())
    assert q.validate() is False


def test_quote_validate_fails_no_line_items() -> None:
    assert _make_quote(line_items=()).validate() is False


def test_quote_validate_fails_expire_before_effective() -> None:
    q = _make_quote(effective_date=date(2026, 6, 1), expire_date=date(2026, 1, 1))
    assert q.validate() is False


def test_quote_validate_fails_empty_created_by() -> None:
    assert _make_quote(created_by="").validate() is False


def test_quote_validate_fails_negative_payment_terms() -> None:
    assert _make_quote(payment_terms=-1).validate() is False


def test_quote_validate_fails_zero_effective_date() -> None:
    assert _make_quote(effective_date=date.min).validate() is False
