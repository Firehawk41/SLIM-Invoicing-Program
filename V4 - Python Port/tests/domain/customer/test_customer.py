import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.customer.customer_repository import _CustomerRow, _FormCustomerRow
from slim.domain.customer.customer_service import CustomerService


@pytest.fixture
def session() -> Session:
    engine = create_engine("sqlite:///:memory:")
    Base.metadata.create_all(engine)
    with Session(engine) as s:
        s.add_all(
            [
                _CustomerRow(
                    id=1,
                    customer_name="Acme Corp",
                    street_address="123 Main St",
                    city="Springfield",
                    state="IL",
                    postal_code="62701",
                    country="USA",
                ),
                _CustomerRow(
                    id=2,
                    customer_name="Beta Labs",
                    street_address="456 Oak Ave",
                    city="Shelbyville",
                    state="TN",
                    postal_code="37160",
                    country="",
                ),
                _FormCustomerRow(
                    form_customer="acme",
                    form_address="123 main st",
                    form_address_2="",
                    customer_id=1,
                ),
                _FormCustomerRow(
                    form_customer="beta labs",
                    form_address="456 oak ave",
                    form_address_2="suite 100",
                    customer_id=2,
                ),
            ]
        )
        s.commit()
        yield s


@pytest.fixture
def service(session: Session) -> CustomerService:
    return CustomerService(session)


# --- Read tests ---


def test_load_by_id_returns_correct_customer(service: CustomerService) -> None:
    c = service.load_customer(1)
    assert c is not None
    assert c.id == 1
    assert c.name == "Acme Corp"
    assert c.street_address == "123 Main St"
    assert c.city == "Springfield"
    assert c.state == "IL"
    assert c.postal_code == "62701"
    assert c.country == "USA"


def test_load_by_id_returns_none_for_missing(service: CustomerService) -> None:
    assert service.load_customer(999) is None


def test_get_by_name_exact_match(service: CustomerService) -> None:
    c = service.get_by_name("Beta Labs")
    assert c is not None
    assert c.id == 2


def test_get_by_name_case_insensitive(service: CustomerService) -> None:
    c = service.get_by_name("beta labs")
    assert c is not None
    assert c.id == 2


def test_get_by_name_returns_none_for_missing(service: CustomerService) -> None:
    assert service.get_by_name("Gamma Corp") is None


def test_all_customers_returns_all_rows(service: CustomerService) -> None:
    assert len(service.all_customers()) == 2


def test_customer_is_frozen(service: CustomerService) -> None:
    c = service.load_customer(1)
    assert c is not None
    with pytest.raises(Exception):
        c.name = "Changed"  # type: ignore[misc]


def test_get_name_by_id_returns_name(service: CustomerService) -> None:
    assert service.get_name_by_id(1) == "Acme Corp"


def test_get_name_by_id_raises_for_missing(service: CustomerService) -> None:
    with pytest.raises(KeyError):
        service.get_name_by_id(999)


# --- FormattedAddress tests ---


def test_formatted_address_with_postal_and_country(service: CustomerService) -> None:
    c = service.load_customer(1)
    assert c is not None
    assert c.formatted_address == "123 Main St, Springfield 62701, USA"


def test_formatted_address_empty_country(service: CustomerService) -> None:
    # Beta Labs has no country
    c = service.load_customer(2)
    assert c is not None
    assert c.formatted_address == "456 Oak Ave, Shelbyville 37160"


def test_formatted_address_state_not_included(service: CustomerService) -> None:
    c = service.load_customer(1)
    assert c is not None
    assert "IL" not in c.formatted_address


# --- Form customer lookup tests ---


def test_get_id_by_form_customer_returns_id(service: CustomerService) -> None:
    assert service.get_id_by_form_customer("acme", "123 main st", "") == 1


def test_get_id_by_form_customer_case_insensitive(service: CustomerService) -> None:
    assert service.get_id_by_form_customer("ACME", "123 Main St", "") == 1


def test_get_id_by_form_customer_returns_zero_for_missing(service: CustomerService) -> None:
    assert service.get_id_by_form_customer("unknown", "nowhere", "") == 0


def test_form_key_exists_true(service: CustomerService) -> None:
    assert service.form_key_exists("beta labs", "456 oak ave", "suite 100") is True


def test_form_key_exists_false(service: CustomerService) -> None:
    assert service.form_key_exists("nobody", "noplace", "") is False


# --- Create tests ---


def test_create_customer_returns_persisted_entity(service: CustomerService) -> None:
    c = service.create_customer(
        name="Gamma Inc",
        street_address="789 Pine Rd",
        city="Capital City",
        state="NY",
        postal_code="10001",
        country="USA",
    )
    assert c.id > 0
    assert c.name == "Gamma Inc"


def test_create_customer_is_persisted_to_db(session: Session) -> None:
    CustomerService(session).create_customer(
        name="Gamma Inc",
        street_address="789 Pine Rd",
        city="Capital City",
        state="NY",
        postal_code="10001",
        country="USA",
    )
    assert CustomerService(session).get_by_name("Gamma Inc") is not None


def test_create_customer_duplicate_name_raises(service: CustomerService) -> None:
    with pytest.raises(ValueError, match="already exists"):
        service.create_customer("Acme Corp", "1 Other St", "Elsewhere", "CA", "90210", "USA")


def test_create_customer_blank_name_raises(service: CustomerService) -> None:
    with pytest.raises(ValueError, match="cannot be blank"):
        service.create_customer("  ", "1 St", "City", "CA", "90210", "USA")


def test_create_customer_blank_address_raises(service: CustomerService) -> None:
    with pytest.raises(ValueError, match="Street address cannot be blank"):
        service.create_customer("Delta Co", "  ", "City", "CA", "90210", "USA")


def test_create_customer_blank_city_raises(service: CustomerService) -> None:
    with pytest.raises(ValueError, match="City cannot be blank"):
        service.create_customer("Delta Co", "1 St", "  ", "CA", "90210", "USA")


# --- Update tests ---


def test_update_customer_persists_changes(session: Session) -> None:
    svc = CustomerService(session)
    original = svc.load_customer(1)
    assert original is not None
    updated = original.model_copy(update={"name": "Acme Corporation", "city": "New City"})
    svc.update_customer(updated)

    fresh = CustomerService(session)
    c = fresh.load_customer(1)
    assert c is not None
    assert c.name == "Acme Corporation"
    assert c.city == "New City"


def test_update_customer_cache_reflects_rename(service: CustomerService) -> None:
    original = service.load_customer(1)
    assert original is not None
    updated = original.model_copy(update={"name": "Acme Corporation"})
    service.update_customer(updated)

    assert service.get_by_name("Acme Corporation") is not None
    assert service.get_by_name("Acme Corp") is None


def test_update_customer_same_name_allowed(service: CustomerService) -> None:
    original = service.load_customer(1)
    assert original is not None
    updated = original.model_copy(update={"city": "New Springfield"})
    result = service.update_customer(updated)
    assert result.city == "New Springfield"


def test_update_customer_duplicate_name_raises(service: CustomerService) -> None:
    original = service.load_customer(1)
    assert original is not None
    conflicting = original.model_copy(update={"name": "Beta Labs"})
    with pytest.raises(ValueError, match="already exists"):
        service.update_customer(conflicting)


def test_update_customer_blank_field_raises(service: CustomerService) -> None:
    original = service.load_customer(1)
    assert original is not None
    blank_city = original.model_copy(update={"city": "  "})
    with pytest.raises(ValueError, match="City cannot be blank"):
        service.update_customer(blank_city)


# --- Delete tests ---


def test_delete_customer_removes_from_cache(service: CustomerService) -> None:
    service.delete_customer(1)
    assert service.load_customer(1) is None
    assert service.get_by_name("Acme Corp") is None


def test_delete_customer_removes_from_db(session: Session) -> None:
    CustomerService(session).delete_customer(1)
    assert CustomerService(session).load_customer(1) is None


def test_delete_nonexistent_customer_is_noop(service: CustomerService) -> None:
    service.delete_customer(999)  # must not raise


# --- add_form_key tests ---


def test_add_form_key_persists_and_indexes(session: Session) -> None:
    svc = CustomerService(session)
    svc.add_form_key("new co", "789 pine rd", "", 1)

    assert svc.form_key_exists("new co", "789 pine rd", "")
    assert svc.get_id_by_form_customer("new co", "789 pine rd", "") == 1

    fresh = CustomerService(session)
    assert fresh.form_key_exists("new co", "789 pine rd", "")


def test_add_form_key_is_idempotent(service: CustomerService) -> None:
    service.add_form_key("acme", "123 main st", "", 1)  # already seeded
    assert service.get_id_by_form_customer("acme", "123 main st", "") == 1  # unchanged
