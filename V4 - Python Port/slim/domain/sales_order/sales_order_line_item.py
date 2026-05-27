from decimal import Decimal

from pydantic import BaseModel, ConfigDict

from slim.domain.tr.enums import ProcessingTime


class SalesOrderLineItem(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    item: str
    description: str
    quantity: int
    rate: Decimal
    quote_price_id: int
    processing_time: ProcessingTime

    @property
    def total(self) -> Decimal:
        return Decimal(self.quantity) * self.rate

    def validate(self) -> bool:
        if self.quantity <= 0:
            return False
        if self.rate < Decimal(0):
            return False
        if not self.item:
            return False
        return True
