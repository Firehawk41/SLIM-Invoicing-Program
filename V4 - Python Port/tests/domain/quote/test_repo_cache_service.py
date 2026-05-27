from datetime import date
from decimal import Decimal

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_cache import QuoteCache
from slim.domain.quote.quote_line_item import QuoteLineItem
from slim.domain.quote.quote_price import QuotePrice
from slim.domain.quote.quote_repository import QuoteRepository
from slim.domain.quote.quote_service import QuoteService
from slim.domain.tr.enums import ProcessingTime, RequestType


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
def repo(session: Session) -> QuoteRepository:
    return QuoteRepository(session)


@pytest.fixture
def cache(repo: QuoteRepository) -> QuoteCache:
    return QuoteCache(repo)


@pytest.fixture
def svc(repo: QuoteRepository, cache: QuoteCache) -> QuoteService:
    cache.build()
    return QuoteService(repo, cache)


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
# Repository — reads on empty DB
# ---------------------------------------------------------------------------

def test_repo_selects_empty_on_fresh_db(repo: QuoteRepository) -> None:
    assert repo.select_all_quotes() == []
    assert repo.select_all_line_items() == []
    assert repo.select_all_prices() == []
    assert repo.select_all_default_quotes() == []


def test_repo_insert_quote_returns_positive_id(repo: QuoteRepository) -> None:
    new_id = repo.insert_quote(_make_quote())
    assert new_id > 0


def test_repo_insert_quote_persists_header(repo: QuoteRepository) -> None:
    repo.insert_quote(_make_quote(payment_terms=45, created_by="alice"))
    rows = repo.select_all_quotes()
    assert len(rows) == 1
    assert rows[0].payment_value == 45
    assert rows[0].created_by == "alice"


def test_repo_insert_quote_persists_line_item_and_price(repo: QuoteRepository) -> None:
    li = _make_line_item(
        bulk_price_min=10,
        prices=(_make_price(regular_price=Decimal("55.00"), bulk_price=Decimal("45.00")),),
    )
    repo.insert_quote(_make_quote(line_items=(li,)))
    assert len(repo.select_all_line_items()) == 1
    pr = repo.select_all_prices()
    assert len(pr) == 1
    assert Decimal(str(pr[0].regular_price)) == Decimal("55.00")


def test_repo_insert_quote_persists_analysis_and_element_ids(repo: QuoteRepository) -> None:
    li = _make_line_item(analysis_ids=(3, 7), element_ids=(14,))
    repo.insert_quote(_make_quote(line_items=(li,)))
    assert len(repo.select_all_analyses()) == 2
    assert len(repo.select_all_elements()) == 1


def test_repo_insert_quote_persists_customer_link(repo: QuoteRepository) -> None:
    repo.insert_quote(_make_quote(customer_ids=(5, 9)))
    custs = repo.select_all_customers()
    assert {c.customer_id for c in custs} == {5, 9}


# ---------------------------------------------------------------------------
# Cache
# ---------------------------------------------------------------------------

def test_cache_empty_after_build_on_empty_db(cache: QuoteCache) -> None:
    cache.build()
    assert cache.all_customer_quotes() == []
    assert cache.get_default_quotes() == []


def test_cache_builds_customer_quote(repo: QuoteRepository, cache: QuoteCache) -> None:
    repo.insert_quote(_make_quote())
    cache.build()
    assert len(cache.all_customer_quotes()) == 1


def test_cache_preserves_price(repo: QuoteRepository, cache: QuoteCache) -> None:
    li = _make_line_item(prices=(_make_price(regular_price=Decimal("77.50")),))
    repo.insert_quote(_make_quote(line_items=(li,)))
    cache.build()
    cached_li = cache.all_customer_quotes()[0].line_items[0]
    assert cached_li.prices[0].regular_price == Decimal("77.50")


def test_cache_preserves_analysis_and_element_ids(repo: QuoteRepository, cache: QuoteCache) -> None:
    li = _make_line_item(analysis_ids=(3, 7), element_ids=(14,))
    repo.insert_quote(_make_quote(line_items=(li,)))
    cache.build()
    cached_li = cache.all_customer_quotes()[0].line_items[0]
    assert set(cached_li.analysis_ids) == {3, 7}
    assert 14 in cached_li.element_ids


def test_cache_preserves_chemical_id(repo: QuoteRepository, cache: QuoteCache) -> None:
    repo.insert_quote(_make_quote(line_items=(_make_line_item(chemical_id=7),)))
    cache.build()
    assert cache.all_customer_quotes()[0].line_items[0].chemical_id == 7


def test_cache_null_chemical_id_round_trips(repo: QuoteRepository, cache: QuoteCache) -> None:
    repo.insert_quote(_make_quote())
    cache.build()
    assert cache.all_customer_quotes()[0].line_items[0].chemical_id is None


def test_cache_get_by_customer_id(repo: QuoteRepository, cache: QuoteCache) -> None:
    repo.insert_quote(_make_quote(customer_ids=(5,)))
    cache.build()
    result = cache.get_by_customer_id(5)
    assert len(result) == 1
    assert result[0].customer_ids == (5,)


def test_cache_get_by_customer_id_empty_on_miss(cache: QuoteCache) -> None:
    cache.build()
    assert cache.get_by_customer_id(999) == []


def test_cache_exists_by_customer_id(repo: QuoteRepository, cache: QuoteCache) -> None:
    repo.insert_quote(_make_quote(customer_ids=(5,)))
    cache.build()
    assert cache.exists_by_customer_id(5) is True
    assert cache.exists_by_customer_id(999) is False


def test_cache_builds_default_quote(repo: QuoteRepository, cache: QuoteCache) -> None:
    dq = _make_quote(is_default_quote=True, customer_ids=())
    repo.insert_default_quote(dq)
    cache.build()
    defaults = cache.get_default_quotes()
    assert len(defaults) == 1
    assert defaults[0].is_default_quote is True


def test_cache_default_quote_has_line_item(repo: QuoteRepository, cache: QuoteCache) -> None:
    li = _make_line_item(request_type=RequestType.WATER)
    dq = _make_quote(is_default_quote=True, customer_ids=(), line_items=(li,))
    repo.insert_default_quote(dq)
    cache.build()
    assert cache.get_default_quotes()[0].line_items[0].request_type == RequestType.WATER


# ---------------------------------------------------------------------------
# Service
# ---------------------------------------------------------------------------

def test_service_all_quotes_empty_initially(svc: QuoteService) -> None:
    assert svc.all_quotes() == []


def test_service_create_quote(svc: QuoteService) -> None:
    new_id = svc.create_quote(_make_quote())
    assert new_id > 0
    quotes = svc.all_quotes()
    assert len(quotes) == 1
    assert quotes[0].id == new_id


def test_service_create_quote_raises_on_invalid(svc: QuoteService) -> None:
    with pytest.raises(ValueError):
        svc.create_quote(_make_quote(line_items=()))


def test_service_get_by_customer_id(svc: QuoteService) -> None:
    svc.create_quote(_make_quote(customer_ids=(5,)))
    result = svc.get_by_customer_id(5)
    assert len(result) == 1


def test_service_exists_by_customer_id(svc: QuoteService) -> None:
    svc.create_quote(_make_quote(customer_ids=(5,)))
    assert svc.exists_by_customer_id(5) is True
    assert svc.exists_by_customer_id(99) is False


def test_service_save_quote_updates_header(svc: QuoteService) -> None:
    new_id = svc.create_quote(_make_quote(payment_terms=30))
    updated = _make_quote(id=new_id, payment_terms=60)
    svc.save_quote(updated)
    assert svc.all_quotes()[0].payment_terms == 60


def test_service_delete_quote(svc: QuoteService) -> None:
    new_id = svc.create_quote(_make_quote())
    svc.delete_quote(new_id)
    assert svc.all_quotes() == []


def test_service_create_default_quote(svc: QuoteService) -> None:
    dq = _make_quote(is_default_quote=True, customer_ids=())
    new_id = svc.create_default_quote(dq)
    assert new_id > 0
    defaults = svc.all_default_quotes()
    assert len(defaults) == 1
    assert defaults[0].is_default_quote is True


def test_service_delete_default_quote(svc: QuoteService) -> None:
    dq = _make_quote(is_default_quote=True, customer_ids=())
    new_id = svc.create_default_quote(dq)
    svc.delete_default_quote(new_id)
    assert svc.all_default_quotes() == []
