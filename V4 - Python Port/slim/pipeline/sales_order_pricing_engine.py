from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_service import QuoteService
from slim.domain.tr.enums import ProcessingTime
from slim.domain.tr.tr_sample import TRSample
from slim.domain.tr.tr_submission import TRSubmission
from slim.pipeline.priced_group import PricedGroup


@dataclass
class _GroupAccumulator:
    quantity: int
    samples: list[str]
    remaining: set[int]
    processing_time: ProcessingTime
    customer_id: int


class SalesOrderPricingEngine:
    """
    Groups submission samples by analysis set + processing time, matches each group
    against valid quotes (customer first, defaults as fallback), assigns prices,
    and merges groups that share a quote_price_id.
    """

    def __init__(self, quote_svc: QuoteService, default_payment_terms: int) -> None:
        self._quote_svc = quote_svc
        self._default_payment_terms = default_payment_terms

    def calculate_pricing_information(
        self,
        submission: TRSubmission,
        today: date | None = None,
    ) -> list[PricedGroup]:
        today = today or date.today()

        groups = self._group_samples(submission)
        matched: list[PricedGroup] = []

        customer_quotes = self._filter_and_sort(
            self._quote_svc.get_by_customer_id(submission.customer_id), today
        )
        self._assign_prices(customer_quotes, groups, matched)

        default_quotes = self._filter_and_sort(
            self._quote_svc.all_default_quotes(), today
        )
        self._assign_prices(default_quotes, groups, matched)

        if groups:
            raise ValueError(
                f"{len(groups)} pricing group(s) could not be matched to any quote. "
                f"CustomerID={submission.customer_id}. "
                "Check that all analyses have valid quotes."
            )

        return self._merge(matched)

    # --- grouping ---

    def _group_samples(self, submission: TRSubmission) -> dict[str, _GroupAccumulator]:
        groups: dict[str, _GroupAccumulator] = {}
        for sample in submission.samples:
            key = self._group_key(submission.customer_id, sample)
            if key in groups:
                acc = groups[key]
                acc.quantity += 1
                acc.samples.append(sample.sample_name)
            else:
                groups[key] = _GroupAccumulator(
                    quantity=1,
                    samples=[sample.sample_name],
                    remaining=set(sample.analysis_ids),
                    processing_time=sample.processing_time,
                    customer_id=submission.customer_id,
                )
        return groups

    def _group_key(self, customer_id: int, sample: TRSample) -> str:
        sorted_ids = sorted(sample.analysis_ids)
        return f"{customer_id}|{int(sample.processing_time)}|{','.join(str(i) for i in sorted_ids)}"

    # --- filtering / sorting ---

    def _filter_and_sort(self, quotes: list[Quote], today: date) -> list[Quote]:
        valid = [
            q
            for q in quotes
            if q.is_valid and q.effective_date <= today <= q.expire_date
        ]
        return sorted(valid, key=self._max_analysis_count, reverse=True)

    def _max_analysis_count(self, quote: Quote) -> int:
        return max((len(li.analysis_ids) for li in quote.line_items), default=0)

    # --- price assignment ---

    def _assign_prices(
        self,
        quotes: list[Quote],
        remaining: dict[str, _GroupAccumulator],
        matched: list[PricedGroup],
    ) -> None:
        consumed = []
        for key, acc in remaining.items():
            for quote in quotes:
                self._try_apply_quote(quote, acc, matched)
                if not acc.remaining:
                    consumed.append(key)
                    break
        for key in consumed:
            del remaining[key]

    def _try_apply_quote(
        self,
        quote: Quote,
        acc: _GroupAccumulator,
        matched: list[PricedGroup],
    ) -> None:
        for li in quote.line_items:
            if not li.is_valid:
                continue
            li_aids = set(li.analysis_ids)
            if not li_aids.issubset(acc.remaining):
                continue
            price = li.get_price(acc.processing_time)
            if price is None or not price.is_valid:
                continue

            actual_price = (
                price.bulk_price
                if li.bulk_price_min > 0
                and acc.quantity >= li.bulk_price_min
                and price.bulk_price >= Decimal(0)
                else price.regular_price
            )
            payment_terms = (
                quote.payment_terms
                if quote.payment_terms > 0
                else self._default_payment_terms
            )

            matched.append(
                PricedGroup(
                    customer_id=acc.customer_id,
                    processing_time=acc.processing_time,
                    quantity=acc.quantity,
                    samples=tuple(acc.samples),
                    analysis_ids=frozenset(li.analysis_ids),
                    quote_price_id=price.id,
                    price=actual_price,
                    payment_terms=payment_terms,
                )
            )
            acc.remaining -= li_aids

            if not acc.remaining:
                return

    # --- merging ---

    def _merge(self, groups: list[PricedGroup]) -> list[PricedGroup]:
        by_quote_price: dict[int, list[PricedGroup]] = {}
        for g in groups:
            by_quote_price.setdefault(g.quote_price_id, []).append(g)

        result = []
        for same in by_quote_price.values():
            if len(same) == 1:
                result.append(same[0])
            else:
                total_qty = sum(g.quantity for g in same)
                all_samples = tuple(s for g in same for s in g.samples)
                result.append(
                    PricedGroup(
                        customer_id=same[0].customer_id,
                        processing_time=same[0].processing_time,
                        quantity=total_qty,
                        samples=all_samples,
                        analysis_ids=same[0].analysis_ids,
                        quote_price_id=same[0].quote_price_id,
                        price=same[0].price,
                        payment_terms=same[0].payment_terms,
                    )
                )
        return result
