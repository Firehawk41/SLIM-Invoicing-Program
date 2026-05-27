from dataclasses import dataclass
from decimal import Decimal

from slim.domain.tr.enums import ProcessingTime


@dataclass(frozen=True)
class PricedGroup:
    """Pricing engine output. One per matched quote price, before analysis name resolution."""

    customer_id: int
    processing_time: ProcessingTime
    quantity: int
    samples: tuple[str, ...]
    analysis_ids: frozenset[int]
    quote_price_id: int
    price: Decimal
    payment_terms: int
