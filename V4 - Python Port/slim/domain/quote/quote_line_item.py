from pydantic import BaseModel, ConfigDict

from slim.domain.tr.enums import ProcessingTime, RequestType
from slim.domain.quote.quote_price import QuotePrice


class QuoteLineItem(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    quote_id: int
    request_type: RequestType
    bulk_price_min: int
    chemical_id: int | None = None  # None = applies to all chemicals
    is_valid: bool
    analysis_ids: tuple[int, ...] = ()
    element_ids: tuple[int, ...] = ()
    prices: tuple[QuotePrice, ...] = ()

    def get_price(self, processing_time: ProcessingTime) -> QuotePrice | None:
        for p in self.prices:
            if p.processing_time == processing_time:
                return p
        return None
