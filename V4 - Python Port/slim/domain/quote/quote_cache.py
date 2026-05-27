from decimal import Decimal

from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_line_item import QuoteLineItem
from slim.domain.quote.quote_price import QuotePrice
from slim.domain.quote.quote_repository import (
    QuoteRepository,
    _DefaultQuoteAnalysisRow,
    _DefaultQuoteLineItemElementRow,
    _DefaultQuoteLineItemRow,
    _DefaultQuotePriceRow,
    _DefaultQuoteRow,
    _QuoteAnalysisRow,
    _QuoteCustomerRow,
    _QuoteLineItemElementRow,
    _QuoteLineItemRow,
    _QuotePriceRow,
    _QuoteRow,
)
from slim.domain.tr.enums import ProcessingTime, RequestType


class QuoteCache:
    """In-memory store for the full quote graph. Full rebuild on every write."""

    def __init__(self, repo: QuoteRepository) -> None:
        self._repo = repo
        self._by_quote_id: dict[int, Quote] = {}
        self._by_customer_id: dict[int, list[Quote]] = {}
        self._default_quotes: list[Quote] = []

    def build(self) -> None:
        self._by_quote_id = {}
        self._by_customer_id = {}
        self._default_quotes = []
        self._build_customer_quotes()
        self._build_default_quotes()

    # --- public reads ---

    def all_customer_quotes(self) -> list[Quote]:
        return list(self._by_quote_id.values())

    def get_by_customer_id(self, customer_id: int) -> list[Quote]:
        return list(self._by_customer_id.get(customer_id, []))

    def exists_by_customer_id(self, customer_id: int) -> bool:
        return customer_id in self._by_customer_id

    def get_default_quotes(self) -> list[Quote]:
        return list(self._default_quotes)

    # --- build helpers ---

    def _build_customer_quotes(self) -> None:
        price_idx = self._build_price_index(self._repo.select_all_prices())
        analysis_idx = self._build_analysis_index(self._repo.select_all_analyses())
        element_idx = self._build_element_index(self._repo.select_all_elements())
        li_idx = self._build_line_item_index(
            self._repo.select_all_line_items(), price_idx, analysis_idx, element_idx
        )
        customer_idx = self._build_customer_index(self._repo.select_all_customers())

        for row in self._repo.select_all_quotes():
            q = self._build_quote(row, li_idx, customer_idx, is_default=False)
            self._by_quote_id[q.id] = q
            for cid in q.customer_ids:
                self._by_customer_id.setdefault(cid, []).append(q)

    def _build_default_quotes(self) -> None:
        price_idx = self._build_price_index(self._repo.select_all_default_prices())
        analysis_idx = self._build_analysis_index(self._repo.select_all_default_analyses())
        element_idx = self._build_element_index(self._repo.select_all_default_elements())
        li_idx = self._build_line_item_index(
            self._repo.select_all_default_line_items(), price_idx, analysis_idx, element_idx
        )

        for row in self._repo.select_all_default_quotes():
            q = self._build_quote(row, li_idx, {}, is_default=True)
            self._default_quotes.append(q)

    def _build_price_index(self, rows: list) -> dict[int, list[QuotePrice]]:
        index: dict[int, list[QuotePrice]] = {}
        for row in rows:
            price = QuotePrice(
                id=row.id,
                quote_line_item_id=row.quote_line_item_id,
                regular_price=Decimal(str(row.regular_price)),
                bulk_price=Decimal(str(row.bulk_price)),
                processing_time=ProcessingTime(row.processing_time_id),
                is_valid=bool(row.valid),
            )
            index.setdefault(row.quote_line_item_id, []).append(price)
        return index

    def _build_analysis_index(self, rows: list) -> dict[int, list[int]]:
        index: dict[int, list[int]] = {}
        for row in rows:
            index.setdefault(row.quote_line_item_id, []).append(row.analysis_id)
        return index

    def _build_element_index(self, rows: list) -> dict[int, list[int]]:
        index: dict[int, list[int]] = {}
        for row in rows:
            index.setdefault(row.quote_line_item_id, []).append(row.element_id)
        return index

    def _build_line_item_index(
        self,
        rows: list,
        price_idx: dict[int, list[QuotePrice]],
        analysis_idx: dict[int, list[int]],
        element_idx: dict[int, list[int]],
    ) -> dict[int, list[QuoteLineItem]]:
        """Returns dict: quote_id -> list[QuoteLineItem]"""
        by_quote: dict[int, list[QuoteLineItem]] = {}
        for row in rows:
            li = QuoteLineItem(
                id=row.id,
                quote_id=row.quote_id,
                request_type=RequestType(row.request_type_id),
                bulk_price_min=row.bulk_min,
                chemical_id=row.chemical_id,
                is_valid=bool(row.valid),
                prices=tuple(price_idx.get(row.id, [])),
                analysis_ids=tuple(analysis_idx.get(row.id, [])),
                element_ids=tuple(element_idx.get(row.id, [])),
            )
            by_quote.setdefault(row.quote_id, []).append(li)
        return by_quote

    def _build_customer_index(self, rows: list) -> dict[int, list[int]]:
        """Returns dict: quote_id -> list[customer_id]"""
        index: dict[int, list[int]] = {}
        for row in rows:
            index.setdefault(row.quote_id, []).append(row.customer_id)
        return index

    def _build_quote(
        self,
        row: _QuoteRow | _DefaultQuoteRow,
        li_by_quote: dict[int, list[QuoteLineItem]],
        customer_by_quote: dict[int, list[int]],
        is_default: bool,
    ) -> Quote:
        return Quote(
            id=row.id,
            effective_date=row.effective_date,
            expire_date=row.expiry_date,
            payment_type=PaymentType(row.payment_type),
            payment_terms=row.payment_value,
            created_by=row.created_by,
            comments=row.comments or "",
            created_date=row.created_date,
            modified_date=row.modified_date,
            is_valid=bool(row.valid),
            is_default_quote=is_default,
            customer_ids=tuple(customer_by_quote.get(row.id, [])),
            line_items=tuple(li_by_quote.get(row.id, [])),
        )
