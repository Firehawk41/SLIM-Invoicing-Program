from datetime import date
from decimal import Decimal
from unittest.mock import MagicMock

import pytest

from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_line_item import QuoteLineItem
from slim.domain.quote.quote_price import QuotePrice
from slim.domain.tr.enums import ProcessingTime, RequestType
from slim.domain.tr.tr_sample import TRSample
from slim.domain.tr.tr_submission import TRSubmission
from slim.pipeline.priced_group import PricedGroup
from slim.pipeline.sales_order_pricing_engine import SalesOrderPricingEngine


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

TODAY = date(2026, 5, 15)
DEFAULT_TERMS = 30


def _make_price(**kwargs) -> QuotePrice:
    defaults = dict(
        id=1,
        quote_line_item_id=1,
        regular_price=Decimal("100.00"),
        bulk_price=Decimal("-1"),
        processing_time=ProcessingTime.NEXT_DAY,
        is_valid=True,
    )
    return QuotePrice(**(defaults | kwargs))


def _make_line_item(**kwargs) -> QuoteLineItem:
    defaults = dict(
        id=1,
        quote_id=1,
        request_type=RequestType.CHEMICAL,
        bulk_price_min=0,
        chemical_id=None,
        is_valid=True,
        analysis_ids=(1,),
        prices=(_make_price(),),
    )
    return QuoteLineItem(**(defaults | kwargs))


def _make_quote(**kwargs) -> Quote:
    defaults = dict(
        id=1,
        effective_date=date(2026, 1, 1),
        expire_date=date(2026, 12, 31),
        payment_type=PaymentType.PO_NUMBER,
        payment_terms=30,
        created_by="admin",
        created_date=date(2026, 1, 1),
        is_valid=True,
        is_default_quote=False,
        customer_ids=(1,),
        line_items=(_make_line_item(),),
    )
    return Quote(**(defaults | kwargs))


def _make_sample(**kwargs) -> TRSample:
    defaults = dict(
        sample_name="S1",
        form_chemical_name="Water",
        processing_time=ProcessingTime.NEXT_DAY,
        additional_notes="",
        requested_time="",
        chemical_id=1,
        analysis_ids=(1,),
    )
    return TRSample(**(defaults | kwargs))


def _make_submission(**kwargs) -> TRSubmission:
    defaults = dict(
        customer_id=1,
        date_submitted=date(2026, 5, 1),
        date_received=date(2026, 5, 1),
        request_type=RequestType.CHEMICAL,
        customer_contact="Jane",
        customer_phone="555-1234",
        po_information="PO-1",
        credit_card_information="",
        file_name="form.xlsx",
        service_date=date(2026, 5, 2),
        download_date=date(2026, 5, 1),  # type: ignore[arg-type]
        form_customer_name="Acme",
        form_customer_address="123 St",
        form_customer_address_2="",
        samples=(_make_sample(),),
    )
    return TRSubmission(**(defaults | kwargs))


def _make_engine(customer_quotes=None, default_quotes=None) -> SalesOrderPricingEngine:
    svc = MagicMock()
    svc.get_by_customer_id.return_value = customer_quotes or []
    svc.all_default_quotes.return_value = default_quotes or []
    return SalesOrderPricingEngine(svc, DEFAULT_TERMS)


# ---------------------------------------------------------------------------
# Filter and sort
# ---------------------------------------------------------------------------


def test_expired_quote_excluded():
    engine = _make_engine(
        customer_quotes=[
            _make_quote(effective_date=date(2025, 1, 1), expire_date=date(2025, 12, 31))
        ]
    )
    with pytest.raises(ValueError, match="could not be matched"):
        engine.calculate_pricing_information(_make_submission(), today=TODAY)


def test_future_quote_excluded():
    engine = _make_engine(
        customer_quotes=[
            _make_quote(effective_date=date(2027, 1, 1), expire_date=date(2027, 12, 31))
        ]
    )
    with pytest.raises(ValueError, match="could not be matched"):
        engine.calculate_pricing_information(_make_submission(), today=TODAY)


def test_invalid_quote_excluded():
    engine = _make_engine(customer_quotes=[_make_quote(is_valid=False)])
    with pytest.raises(ValueError, match="could not be matched"):
        engine.calculate_pricing_information(_make_submission(), today=TODAY)


def test_bundles_sorted_before_singles():
    """Quote with more analyses per line item should be tried first."""
    single = _make_quote(id=1, line_items=(_make_line_item(id=1, analysis_ids=(1,)),))
    bundle = _make_quote(
        id=2,
        line_items=(
            _make_line_item(
                id=2,
                analysis_ids=(1, 2),
                prices=(
                    _make_price(id=2, processing_time=ProcessingTime.NEXT_DAY),
                ),
            ),
        ),
    )
    engine = _make_engine(customer_quotes=[single, bundle])
    submission = _make_submission(
        samples=(
            _make_sample(analysis_ids=(1, 2)),
        )
    )
    groups = engine.calculate_pricing_information(submission, today=TODAY)
    assert len(groups) == 1
    assert groups[0].quote_price_id == 2


# ---------------------------------------------------------------------------
# Basic matching
# ---------------------------------------------------------------------------


def test_single_analysis_match():
    engine = _make_engine(customer_quotes=[_make_quote()])
    groups = engine.calculate_pricing_information(_make_submission(), today=TODAY)
    assert len(groups) == 1
    assert groups[0].quantity == 1
    assert groups[0].price == Decimal("100.00")
    assert groups[0].analysis_ids == frozenset({1})


def test_unmatched_raises():
    engine = _make_engine()
    with pytest.raises(ValueError, match="could not be matched"):
        engine.calculate_pricing_information(_make_submission(), today=TODAY)


def test_default_quote_fallback():
    engine = _make_engine(
        customer_quotes=[],
        default_quotes=[_make_quote(is_default_quote=True, customer_ids=())],
    )
    groups = engine.calculate_pricing_information(_make_submission(), today=TODAY)
    assert len(groups) == 1


def test_payment_terms_from_quote():
    engine = _make_engine(customer_quotes=[_make_quote(payment_terms=45)])
    groups = engine.calculate_pricing_information(_make_submission(), today=TODAY)
    assert groups[0].payment_terms == 45


def test_payment_terms_defaults_when_zero():
    engine = _make_engine(customer_quotes=[_make_quote(payment_terms=0)])
    groups = engine.calculate_pricing_information(_make_submission(), today=TODAY)
    assert groups[0].payment_terms == DEFAULT_TERMS


# ---------------------------------------------------------------------------
# Bulk pricing
# ---------------------------------------------------------------------------


def test_bulk_price_applied_when_quantity_qualifies():
    price = _make_price(
        id=1,
        regular_price=Decimal("100.00"),
        bulk_price=Decimal("75.00"),
        processing_time=ProcessingTime.NEXT_DAY,
    )
    li = _make_line_item(bulk_price_min=3, prices=(price,))
    engine = _make_engine(customer_quotes=[_make_quote(line_items=(li,))])

    submission = _make_submission(
        samples=tuple(_make_sample(sample_name=f"S{i}") for i in range(5))
    )
    groups = engine.calculate_pricing_information(submission, today=TODAY)
    assert groups[0].price == Decimal("75.00")


def test_regular_price_when_below_bulk_min():
    price = _make_price(
        id=1,
        regular_price=Decimal("100.00"),
        bulk_price=Decimal("75.00"),
        processing_time=ProcessingTime.NEXT_DAY,
    )
    li = _make_line_item(bulk_price_min=10, prices=(price,))
    engine = _make_engine(customer_quotes=[_make_quote(line_items=(li,))])
    groups = engine.calculate_pricing_information(_make_submission(), today=TODAY)
    assert groups[0].price == Decimal("100.00")


# ---------------------------------------------------------------------------
# Bundle splitting (one group matched across multiple line items)
# ---------------------------------------------------------------------------


def test_bundle_split_across_two_line_items():
    """A group with analyses {1, 2} is split into two PricedGroups."""
    li1 = _make_line_item(id=1, analysis_ids=(1,), prices=(_make_price(id=1),))
    li2 = _make_line_item(
        id=2,
        analysis_ids=(2,),
        prices=(_make_price(id=2, processing_time=ProcessingTime.NEXT_DAY),),
    )
    engine = _make_engine(customer_quotes=[_make_quote(line_items=(li1, li2))])
    submission = _make_submission(samples=(_make_sample(analysis_ids=(1, 2)),))
    groups = engine.calculate_pricing_information(submission, today=TODAY)
    assert len(groups) == 2
    qp_ids = {g.quote_price_id for g in groups}
    assert qp_ids == {1, 2}


# ---------------------------------------------------------------------------
# Merging
# ---------------------------------------------------------------------------


def test_identical_quote_price_groups_are_merged():
    """Two samples with the same analysis set → one PricedGroup after merge."""
    engine = _make_engine(customer_quotes=[_make_quote()])
    submission = _make_submission(
        samples=(
            _make_sample(sample_name="S1"),
            _make_sample(sample_name="S2"),
        )
    )
    groups = engine.calculate_pricing_information(submission, today=TODAY)
    assert len(groups) == 1
    assert groups[0].quantity == 2
    assert set(groups[0].samples) == {"S1", "S2"}


def test_different_processing_times_not_merged():
    price_nd = _make_price(id=1, processing_time=ProcessingTime.NEXT_DAY)
    price_sd = _make_price(id=2, processing_time=ProcessingTime.SAME_DAY_RUSH)
    li = _make_line_item(prices=(price_nd, price_sd))
    engine = _make_engine(customer_quotes=[_make_quote(line_items=(li,))])
    submission = _make_submission(
        samples=(
            _make_sample(sample_name="S1", processing_time=ProcessingTime.NEXT_DAY),
            _make_sample(sample_name="S2", processing_time=ProcessingTime.SAME_DAY_RUSH),
        )
    )
    groups = engine.calculate_pricing_information(submission, today=TODAY)
    assert len(groups) == 2


# ---------------------------------------------------------------------------
# Date injection
# ---------------------------------------------------------------------------


def test_today_injection_controls_validity():
    quote = _make_quote(
        effective_date=date(2026, 6, 1), expire_date=date(2026, 6, 30)
    )
    engine = _make_engine(customer_quotes=[quote])

    # Before effective date — should not match
    with pytest.raises(ValueError, match="could not be matched"):
        engine.calculate_pricing_information(_make_submission(), today=date(2026, 5, 31))

    # Within range — should match
    groups = engine.calculate_pricing_information(_make_submission(), today=date(2026, 6, 15))
    assert len(groups) == 1
