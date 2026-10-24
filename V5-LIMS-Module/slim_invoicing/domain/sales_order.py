"""A sales order: what the lab invoices for one TR (one per TR, D-83)."""

from datetime import date, datetime
from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field

from slim_invoicing.domain.enums import PaymentMethod, ProcessingTime


class _Frozen(BaseModel):
    model_config = ConfigDict(frozen=True)


class SalesOrderLineItem(_Frozen):
    item: str = Field(min_length=1)
    description: str
    quantity: int = Field(ge=1)
    rate: Decimal = Field(ge=0)
    quote_price_id: int
    processing_time: ProcessingTime
    samples: tuple[str, ...]

    @property
    def total(self) -> Decimal:
        return self.quantity * self.rate


class SalesOrder(_Frozen):
    id: int = 0
    tr_number: str = Field(min_length=1)
    customer_id: int
    customer_name: str = Field(min_length=1)
    payment_method: PaymentMethod
    payment_terms: int = Field(ge=0)
    po_number: str = ""
    requester_name: str = ""
    requester_email: str = ""
    date_received: date
    service_date: date
    line_items: tuple[SalesOrderLineItem, ...] = Field(min_length=1)
    created_by: str = Field(min_length=1)
    created_at: datetime | None = None  # database clock

    @property
    def total(self) -> Decimal:
        return sum((li.total for li in self.line_items), Decimal(0))
