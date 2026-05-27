from datetime import date

from pydantic import BaseModel, ConfigDict

from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote_line_item import QuoteLineItem


class Quote(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    effective_date: date
    expire_date: date
    payment_type: PaymentType
    payment_terms: int
    created_by: str
    comments: str = ""
    created_date: date
    modified_date: date | None = None
    is_valid: bool
    is_default_quote: bool
    customer_ids: tuple[int, ...] = ()
    line_items: tuple[QuoteLineItem, ...] = ()

    def validate(self) -> bool:
        if self.id < 0:
            return False
        if not self.is_default_quote and not self.customer_ids:
            return False
        if not self.line_items:
            return False
        if self.payment_terms < 0:
            return False
        if not self.created_by:
            return False
        if self.effective_date == date.min:
            return False
        if self.expire_date == date.min:
            return False
        if self.created_date == date.min:
            return False
        if self.expire_date < self.effective_date:
            return False
        return True
