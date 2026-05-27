from sqlalchemy import ForeignKey, String
from sqlalchemy.orm import Mapped, Session, mapped_column

from infrastructure.database import Base
from slim.domain.customer.customer import Customer


class _CustomerRow(Base):
    __tablename__ = "customers"

    id: Mapped[int] = mapped_column("ID", primary_key=True)
    customer_name: Mapped[str] = mapped_column(String)
    street_address: Mapped[str] = mapped_column(String)
    city: Mapped[str] = mapped_column(String)
    state: Mapped[str] = mapped_column(String)
    postal_code: Mapped[str] = mapped_column(String)
    country: Mapped[str] = mapped_column(String)


class _FormCustomerRow(Base):
    __tablename__ = "form_customers"

    form_customer: Mapped[str] = mapped_column(String, primary_key=True)
    form_address: Mapped[str] = mapped_column(String, primary_key=True)
    form_address_2: Mapped[str] = mapped_column(String, primary_key=True)
    customer_id: Mapped[int] = mapped_column(ForeignKey("customers.ID"))


class CustomerRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def select_all(self) -> list[Customer]:
        """Load all customers ordered by name. Called only by CustomerCache."""
        rows = (
            self._session.query(_CustomerRow)
            .order_by(_CustomerRow.customer_name)
            .all()
        )
        return [_to_customer(r) for r in rows]

    def select_all_form_customers(self) -> list[tuple[str, str, str, int]]:
        """Load all form-customer mappings. Called only by CustomerCache."""
        rows = (
            self._session.query(_FormCustomerRow)
            .order_by(_FormCustomerRow.form_customer)
            .all()
        )
        return [(r.form_customer, r.form_address, r.form_address_2, r.customer_id) for r in rows]

    def insert(
        self,
        name: str,
        street_address: str,
        city: str,
        state: str,
        postal_code: str,
        country: str,
    ) -> Customer:
        """Insert a new customer. Returns the persisted Customer with its new ID."""
        row = _CustomerRow(
            customer_name=name,
            street_address=street_address,
            city=city,
            state=state,
            postal_code=postal_code,
            country=country,
        )
        self._session.add(row)
        self._session.flush()  # materialise the auto-assigned ID
        self._session.commit()
        return Customer(
            id=row.id,
            name=name,
            street_address=street_address,
            city=city,
            state=state,
            postal_code=postal_code,
            country=country,
        )

    def update(self, customer: Customer) -> None:
        """Overwrite all mutable fields for an existing customer."""
        row = self._session.get(_CustomerRow, customer.id)
        if row is None:
            raise ValueError(f"Customer ID not found: {customer.id}")
        row.customer_name = customer.name
        row.street_address = customer.street_address
        row.city = customer.city
        row.state = customer.state
        row.postal_code = customer.postal_code
        row.country = customer.country
        self._session.commit()

    def delete(self, customer_id: int) -> None:
        """Delete the customer row. No-op if the ID does not exist."""
        row = self._session.get(_CustomerRow, customer_id)
        if row is not None:
            self._session.delete(row)
            self._session.commit()

    def insert_form_customer(
        self,
        form_customer: str,
        form_address: str,
        form_address_2: str,
        customer_id: int,
    ) -> None:
        """Persist a new form-customer -> customer_id mapping."""
        self._session.add(
            _FormCustomerRow(
                form_customer=form_customer,
                form_address=form_address,
                form_address_2=form_address_2,
                customer_id=customer_id,
            )
        )
        self._session.commit()


def _to_customer(row: _CustomerRow) -> Customer:
    return Customer(
        id=row.id,
        name=row.customer_name,
        street_address=row.street_address,
        city=row.city,
        state=row.state,
        postal_code=row.postal_code,
        country=row.country,
    )
