from sqlalchemy.orm import Session

from slim.domain.customer.customer import Customer
from slim.domain.customer.customer_cache import CustomerCache
from slim.domain.customer.customer_repository import CustomerRepository


class CustomerService:
    """Public API for the customer domain. Callers never touch repo or cache."""

    def __init__(self, session: Session) -> None:
        self._repo = CustomerRepository(session)
        self._cache = CustomerCache(self._repo)

    # --- Reads ---

    def load_customer(self, customer_id: int) -> Customer | None:
        """Return the Customer for customer_id, or None if not found."""
        return self._cache.get_by_id(customer_id)

    def get_by_name(self, name: str) -> Customer | None:
        """Return the Customer matching name (case-insensitive), or None if not found."""
        return self._cache.get_by_name(name)

    def all_customers(self) -> list[Customer]:
        """Return all customers loaded at startup."""
        return self._cache.all_customers()

    def get_id_by_form_customer(
        self, form_name: str, form_address: str, form_address_2: str
    ) -> int:
        """Return the customer ID for the composite form key, or 0 if not mapped."""
        return self._cache.get_id_by_form_customer(form_name, form_address, form_address_2)

    def form_key_exists(self, name: str, address: str, address_2: str) -> bool:
        """Return True if the composite form key is already mapped."""
        return self._cache.form_key_exists(name, address, address_2)

    def get_name_by_id(self, customer_id: int) -> str:
        """Return the customer name for customer_id. Raises KeyError if not found."""
        return self._cache.get_name_by_id(customer_id)

    # --- Writes ---

    def create_customer(
        self,
        name: str,
        street_address: str,
        city: str,
        state: str,
        postal_code: str,
        country: str,
    ) -> Customer:
        """Validate, persist, and cache a new customer. Returns the persisted entity."""
        self._require_valid_name(name)
        self._require_unique_name(name)
        self._require_non_blank(street_address, "Street address")
        self._require_non_blank(city, "City")
        self._require_non_blank(state, "State")
        self._require_non_blank(postal_code, "Postal code")
        self._require_non_blank(country, "Country")

        customer = self._repo.insert(name, street_address, city, state, postal_code, country)
        self._cache.upsert(customer)
        return customer

    def update_customer(self, customer: Customer) -> Customer:
        """Validate and persist changes to an existing customer. Returns the same entity."""
        if customer.id <= 0:
            raise ValueError("Cannot update a customer with no ID")
        self._require_valid_name(customer.name)
        self._require_unique_name_for_update(customer.name, customer.id)
        self._require_non_blank(customer.street_address, "Street address")
        self._require_non_blank(customer.city, "City")
        self._require_non_blank(customer.state, "State")
        self._require_non_blank(customer.postal_code, "Postal code")
        self._require_non_blank(customer.country, "Country")

        self._repo.update(customer)
        self._cache.upsert(customer)
        return customer

    def delete_customer(self, customer_id: int) -> None:
        """Delete the customer from DB and cache."""
        self._repo.delete(customer_id)
        self._cache.remove(customer_id)

    def add_form_key(
        self,
        form_customer: str,
        form_address: str,
        form_address_2: str,
        customer_id: int,
    ) -> None:
        """Persist a form-key -> customer_id mapping. Idempotent: no-op if already mapped."""
        if self._cache.form_key_exists(form_customer, form_address, form_address_2):
            return
        self._repo.insert_form_customer(form_customer, form_address, form_address_2, customer_id)
        self._cache.add_form_customer_index(form_customer, form_address, form_address_2, customer_id)

    # --- Validation (public for UI pre-validation) ---

    def is_valid_name(self, name: str) -> bool:
        return len(name.strip()) > 0

    def is_unique_name(self, name: str) -> bool:
        return not self._cache.exists_by_name(name)

    def is_unique_name_for_update(self, name: str, customer_id: int) -> bool:
        """Return True if name is available, treating the customer's own name as non-conflicting."""
        existing = self._cache.get_by_name(name)
        return existing is None or existing.id == customer_id

    # --- Private helpers ---

    def _require_valid_name(self, name: str) -> None:
        if not self.is_valid_name(name):
            raise ValueError("Customer name cannot be blank")

    def _require_unique_name(self, name: str) -> None:
        if not self.is_unique_name(name):
            raise ValueError(f"Customer name already exists: {name}")

    def _require_unique_name_for_update(self, name: str, customer_id: int) -> None:
        if not self.is_unique_name_for_update(name, customer_id):
            raise ValueError(f"Customer name already exists: {name}")

    def _require_non_blank(self, value: str, field_name: str) -> None:
        if not value.strip():
            raise ValueError(f"{field_name} cannot be blank")
