from datetime import date
from decimal import Decimal

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem
from slim.domain.sales_order.sales_order_service import SalesOrderService
from slim.domain.tr.enums import ProcessingTime


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
def svc(session: Session) -> SalesOrderService:
    return SalesOrderService(session)


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


def _make_so(**kwargs) -> SalesOrder:
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
        line_items=(_make_line_item(),),
    )
    return SalesOrder(**(defaults | kwargs))


# ---------------------------------------------------------------------------
# create
# ---------------------------------------------------------------------------


def test_create_returns_sealed_so(svc: SalesOrderService):
    so = svc.create(_make_so())
    assert so.id > 0


def test_create_seals_line_item_ids(svc: SalesOrderService):
    so = svc.create(_make_so())
    assert all(li.id > 0 for li in so.line_items)


def test_create_round_trip(svc: SalesOrderService):
    original = _make_so(customer_name="Round Trip Co")
    created = svc.create(original)
    loaded = svc.load(created.id)
    assert loaded is not None
    assert loaded.customer_name == "Round Trip Co"
    assert loaded.id == created.id


def test_create_line_items_persisted(svc: SalesOrderService):
    original = _make_so(
        line_items=(
            _make_line_item(item="ICP-MS - Next Day", quantity=3),
            _make_line_item(item="Si - Next Day", quantity=2),
        )
    )
    created = svc.create(original)
    loaded = svc.load(created.id)
    assert loaded is not None
    assert len(loaded.line_items) == 2
    items = {li.item for li in loaded.line_items}
    assert items == {"ICP-MS - Next Day", "Si - Next Day"}


def test_create_validation_failure_raises(svc: SalesOrderService):
    with pytest.raises(ValueError, match="validation"):
        svc.create(_make_so(customer_id=0))


# ---------------------------------------------------------------------------
# load
# ---------------------------------------------------------------------------


def test_load_not_found_returns_none(svc: SalesOrderService):
    assert svc.load(9999) is None


def test_load_preserves_all_fields(svc: SalesOrderService):
    original = _make_so(
        customer_name="Field Check",
        submission_ref="test.xlsx",
        po_number="PO-1234",
        payment_type=PaymentType.PO_NUMBER,
        payment_terms=45,
        requester_name="Bob",
        requester_email="bob@lab.com",
    )
    created = svc.create(original)
    loaded = svc.load(created.id)
    assert loaded is not None
    assert loaded.customer_name == "Field Check"
    assert loaded.submission_ref == "test.xlsx"
    assert loaded.po_number == "PO-1234"
    assert loaded.payment_terms == 45
    assert loaded.requester_name == "Bob"
    assert loaded.requester_email == "bob@lab.com"


# ---------------------------------------------------------------------------
# load_by_date_range
# ---------------------------------------------------------------------------


def test_load_by_date_range_returns_matching(svc: SalesOrderService):
    svc.create(_make_so(service_date=date(2026, 3, 1)))
    svc.create(_make_so(service_date=date(2026, 4, 15)))
    svc.create(_make_so(service_date=date(2026, 5, 1)))

    results = svc.load_by_date_range(date(2026, 4, 1), date(2026, 4, 30))
    assert len(results) == 1
    assert results[0].service_date == date(2026, 4, 15)


def test_load_by_date_range_inclusive_boundaries(svc: SalesOrderService):
    svc.create(_make_so(service_date=date(2026, 4, 1)))
    svc.create(_make_so(service_date=date(2026, 4, 30)))

    results = svc.load_by_date_range(date(2026, 4, 1), date(2026, 4, 30))
    assert len(results) == 2


# ---------------------------------------------------------------------------
# save
# ---------------------------------------------------------------------------


def test_save_updates_header(svc: SalesOrderService):
    created = svc.create(_make_so(customer_name="Original"))
    updated = created.model_copy(update={"customer_name": "Updated"})
    svc.save(updated)
    loaded = svc.load(created.id)
    assert loaded is not None
    assert loaded.customer_name == "Updated"


def test_save_validation_failure_raises(svc: SalesOrderService):
    created = svc.create(_make_so())
    invalid = created.model_copy(update={"customer_id": 0})
    with pytest.raises(ValueError, match="validation"):
        svc.save(invalid)


# ---------------------------------------------------------------------------
# delete
# ---------------------------------------------------------------------------


def test_delete_removes_so(svc: SalesOrderService):
    created = svc.create(_make_so())
    svc.delete(created.id)
    assert svc.load(created.id) is None


def test_delete_cascades_to_line_items(svc: SalesOrderService, session: Session):
    from sqlalchemy import text

    created = svc.create(_make_so(line_items=(_make_line_item(),)))
    svc.delete(created.id)
    count = session.execute(
        text("SELECT COUNT(*) FROM sales_order_items WHERE sales_order_id=:sid"),
        {"sid": created.id},
    ).scalar()
    assert count == 0
