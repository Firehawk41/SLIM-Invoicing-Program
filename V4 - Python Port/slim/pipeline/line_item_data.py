from dataclasses import dataclass
from decimal import Decimal

from slim.domain.tr.enums import ProcessingTime


@dataclass(frozen=True)
class LineItemData:
    """Builder's enriched DTO. Carries resolved item/description strings for entity construction."""

    item: str
    description: str
    quantity: int
    price: Decimal
    quote_price_id: int
    processing_time: ProcessingTime
    payment_terms: int
