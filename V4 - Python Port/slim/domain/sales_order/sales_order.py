from datetime import date
from decimal import Decimal

from pydantic import BaseModel, ConfigDict

from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem


class SalesOrder(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    customer_name: str
    customer_id: int
    submission_id: int
    submission_ref: str
    date_received: date
    po_number: str
    payment_type: PaymentType
    payment_terms: int
    requester_name: str
    requester_email: str
    service_date: date
    line_items: tuple[SalesOrderLineItem, ...] = ()

    @property
    def total_value(self) -> Decimal:
        return sum((li.total for li in self.line_items), Decimal(0))

    def validate(self) -> bool:
        if self.customer_id == 0:
            return False
        if not self.customer_name:
            return False
        if self.service_date == date.min:
            return False
        if self.date_received == date.min:
            return False
        if self.payment_type == PaymentType.PO_NUMBER and not self.po_number:
            return False
        if self.payment_terms < 0:
            return False
        return True
