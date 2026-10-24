"""Quotes: what a customer (or everyone, for a default quote) pays per analysis set and turnaround.

Ported from V4 unchanged in meaning. IDs of customers and analyses are slim-domain's (D-83).
"""

from datetime import date
from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from slim_invoicing.domain.enums import PaymentMethod, ProcessingTime


class _Frozen(BaseModel):
    model_config = ConfigDict(frozen=True)


class QuotePrice(_Frozen):
    id: int = 0
    processing_time: ProcessingTime
    regular_price: Decimal = Field(ge=0)
    bulk_price: Decimal | None = Field(default=None, ge=0)  # None: no bulk price


class QuoteLineItem(_Frozen):
    id: int = 0
    analysis_ids: tuple[int, ...] = Field(min_length=1)  # priced together, as one line
    bulk_price_min: int = Field(default=0, ge=0)  # 0: no bulk pricing
    prices: tuple[QuotePrice, ...] = Field(min_length=1)

    def price_for(self, processing_time: ProcessingTime) -> QuotePrice | None:
        return next((p for p in self.prices if p.processing_time == processing_time), None)


class Quote(_Frozen):
    id: int = 0
    effective_date: date
    expire_date: date
    payment_method: PaymentMethod = PaymentMethod.PURCHASE_ORDER
    payment_terms: int = Field(default=0, ge=0)  # days; 0: the lab's default
    is_default: bool = False  # applies to every customer, after their own quotes
    customer_ids: tuple[int, ...] = ()
    line_items: tuple[QuoteLineItem, ...] = Field(min_length=1)
    created_by: str = Field(min_length=1)
    comments: str = ""

    @model_validator(mode="after")
    def _check(self) -> "Quote":
        if self.expire_date < self.effective_date:
            raise ValueError("a quote can't expire before it takes effect")
        if not self.is_default and not self.customer_ids:
            raise ValueError("a customer quote needs at least one customer")
        return self

    def valid_on(self, day: date) -> bool:
        return self.effective_date <= day <= self.expire_date

    @property
    def widest_line(self) -> int:
        return max(len(li.analysis_ids) for li in self.line_items)
