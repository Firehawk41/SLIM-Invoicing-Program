from decimal import Decimal

from pydantic import BaseModel, ConfigDict

from slim.domain.tr.enums import ProcessingTime

SENTINEL_BULK_PRICE = Decimal("-1")


class QuotePrice(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    quote_line_item_id: int
    regular_price: Decimal
    bulk_price: Decimal  # Decimal("-1") = no bulk pricing available
    processing_time: ProcessingTime
    is_valid: bool

    @property
    def has_bulk_pricing(self) -> bool:
        return self.bulk_price >= 0
