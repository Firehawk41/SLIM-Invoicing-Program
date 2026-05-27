from datetime import date
from decimal import Decimal
from unittest.mock import MagicMock

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.analysis.analysis import Analysis
from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_service import SalesOrderService
from slim.domain.tr.enums import ProcessingTime, RequestType
from slim.domain.tr.tr_sample import TRSample
from slim.domain.tr.tr_submission import TRSubmission
from slim.pipeline.line_item_data import LineItemData
from slim.pipeline.priced_group import PricedGroup
from slim.pipeline.sales_order_line_item_builder import SalesOrderLineItemBuilder


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def session() -> Session:
    engine = create_engine("sqlite:///:memory:")
    Base.metadata.create_all(engine)
    with Session(engine) as s:
        yield s


@pytest.fixture
def so_svc(session: Session) -> SalesOrderService:
    return SalesOrderService(session)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_analysis(id: int, name: str, description: str) -> Analysis:
    return Analysis(id=id, name=name, description=description)


def _make_priced_group(**kwargs) -> PricedGroup:
    defaults = dict(
        customer_id=1,
        processing_time=ProcessingTime.NEXT_DAY,
        quantity=5,
        samples=("S1", "S2", "S3", "S4", "S5"),
        analysis_ids=frozenset({1}),
        quote_price_id=10,
        price=Decimal("100.00"),
        payment_terms=30,
    )
    return PricedGroup(**(defaults | kwargs))


def _make_analysis_svc(analyses: dict[int, Analysis]) -> MagicMock:
    svc = MagicMock()
    svc.load_analysis.side_effect = lambda aid: analyses.get(aid)
    return svc


def _make_line_item_builder(
    analyses: dict[int, Analysis] | None = None,
    priced_groups: list[PricedGroup] | None = None,
) -> SalesOrderLineItemBuilder:
    if analyses is None:
        analyses = {1: _make_analysis(1, "ICP-MS", "36 trace elements by ICP-MS")}
    pricing_engine = MagicMock()
    pricing_engine.calculate_pricing_information.return_value = priced_groups or []
    return SalesOrderLineItemBuilder(pricing_engine, _make_analysis_svc(analyses))


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
        samples=(
            TRSample(
                sample_name="S1",
                form_chemical_name="Water",
                processing_time=ProcessingTime.NEXT_DAY,
                additional_notes="",
                requested_time="",
                chemical_id=1,
                analysis_ids=(1,),
            ),
        ),
    )
    return TRSubmission(**(defaults | kwargs))


# ---------------------------------------------------------------------------
# SalesOrderLineItemBuilder.build_from_priced_groups
# ---------------------------------------------------------------------------


def test_build_from_priced_groups_happy_path():
    analyses = {1: _make_analysis(1, "ICP-MS", "36 trace elements by ICP-MS")}
    builder = _make_line_item_builder(analyses=analyses)
    group = _make_priced_group()

    items = builder.build_from_priced_groups([group])

    assert len(items) == 1
    assert items[0].item == "ICP-MS - Next Day"
    assert items[0].description == "36 trace elements by ICP-MS - Next Day"
    assert items[0].quantity == 5
    assert items[0].rate == Decimal("100.00")
    assert items[0].quote_price_id == 10


def test_build_item_string_multi_analysis():
    analyses = {
        1: _make_analysis(1, "ICP-MS", "36 trace elements"),
        2: _make_analysis(2, "Si", "Silicon by ICP-MS"),
    }
    builder = _make_line_item_builder(analyses=analyses)
    group = _make_priced_group(analysis_ids=frozenset({1, 2}))

    items = builder.build_from_priced_groups([group])

    assert items[0].item == "ICP-MS + Si - Next Day"
    assert items[0].description == "36 trace elements + Silicon by ICP-MS - Next Day"


def test_build_analysis_not_found_raises():
    builder = _make_line_item_builder(analyses={})
    with pytest.raises(ValueError, match="Analysis ID not found"):
        builder.build_from_priced_groups([_make_priced_group()])


def test_build_validation_failure_raises():
    analyses = {1: _make_analysis(1, "ICP-MS", "36 trace elements by ICP-MS")}
    builder = _make_line_item_builder(analyses=analyses)
    group = _make_priced_group(quantity=0)
    with pytest.raises(ValueError, match="validation"):
        builder.build_from_priced_groups([group])


def test_build_multiple_groups():
    analyses = {
        1: _make_analysis(1, "ICP-MS", "36 trace elements"),
        2: _make_analysis(2, "Si", "Silicon"),
    }
    builder = _make_line_item_builder(analyses=analyses)
    groups = [
        _make_priced_group(analysis_ids=frozenset({1}), quote_price_id=1),
        _make_priced_group(analysis_ids=frozenset({2}), quote_price_id=2),
    ]
    items = builder.build_from_priced_groups(groups)
    assert len(items) == 2


# ---------------------------------------------------------------------------
# SalesOrderBuilder.build_from_submission (integration-ish)
# ---------------------------------------------------------------------------


def test_so_builder_builds_and_persists(so_svc: SalesOrderService):
    from slim.pipeline.sales_order_builder import SalesOrderBuilder

    analyses = {1: _make_analysis(1, "ICP-MS", "36 trace elements by ICP-MS")}
    priced_group = _make_priced_group(payment_terms=45)

    pricing_engine = MagicMock()
    pricing_engine.calculate_pricing_information.return_value = [priced_group]

    li_builder = SalesOrderLineItemBuilder(
        pricing_engine, _make_analysis_svc(analyses)
    )

    customer_svc = MagicMock()
    customer_svc.get_name_by_id.return_value = "Acme Corp"

    builder = SalesOrderBuilder(li_builder, customer_svc, so_svc)
    so = builder.build_from_submission(_make_submission())

    assert so.id > 0
    assert so.customer_name == "Acme Corp"
    assert so.payment_terms == 45
    assert len(so.line_items) == 1
    assert so.line_items[0].item == "ICP-MS - Next Day"


def test_so_builder_payment_type_credit_card(so_svc: SalesOrderService):
    from slim.pipeline.sales_order_builder import SalesOrderBuilder

    analyses = {1: _make_analysis(1, "ICP-MS", "36 trace elements")}
    priced_group = _make_priced_group()

    pricing_engine = MagicMock()
    pricing_engine.calculate_pricing_information.return_value = [priced_group]

    li_builder = SalesOrderLineItemBuilder(
        pricing_engine, _make_analysis_svc(analyses)
    )
    customer_svc = MagicMock()
    customer_svc.get_name_by_id.return_value = "CC Customer"

    builder = SalesOrderBuilder(li_builder, customer_svc, so_svc)
    submission = _make_submission(credit_card_information="4111111111111111")
    so = builder.build_from_submission(submission)

    assert so.payment_type == PaymentType.CREDIT_CARD


def test_so_builder_empty_groups_uses_default_payment_terms(so_svc: SalesOrderService):
    from slim.pipeline.sales_order_builder import SalesOrderBuilder

    pricing_engine = MagicMock()
    pricing_engine.calculate_pricing_information.return_value = []

    li_builder = SalesOrderLineItemBuilder(pricing_engine, MagicMock())
    customer_svc = MagicMock()
    customer_svc.get_name_by_id.return_value = "Acme"

    builder = SalesOrderBuilder(li_builder, customer_svc, so_svc)
    submission = _make_submission(samples=())
    so = builder.build_from_submission(submission)

    assert so.payment_terms == SalesOrderBuilder.DEFAULT_PAYMENT_TERMS
