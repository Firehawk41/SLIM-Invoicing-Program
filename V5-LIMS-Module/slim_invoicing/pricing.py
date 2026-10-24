"""The pricing engine, ported from V4 unchanged in behaviour (D-83).

Samples are grouped by turnaround + analysis set. Each group is matched against the customer's own
quotes first, then the default quotes, each list tried widest line first, so a bundle ("36 Elements +
7 Anions") wins over its parts. A line applies when all its analyses are still unpriced in the group
and it has a price for the group's turnaround; the bulk price applies once the group reaches the
line's minimum. Groups that end on the same quote price are merged into one line. A group left with
an unpriced analysis is an error, never a zero-price line.
"""

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from slim_invoicing.domain.billable import BillableSample
from slim_invoicing.domain.enums import ProcessingTime
from slim_invoicing.domain.quote import Quote


class Unpriced(ValueError):
    """Some analyses match no valid quote: the lab must quote them before invoicing."""


@dataclass(frozen=True)
class PricedGroup:
    processing_time: ProcessingTime
    quantity: int
    samples: tuple[str, ...]
    analysis_ids: frozenset[int]
    quote_price_id: int
    price: Decimal
    payment_terms: int  # 0: the lab's default


@dataclass
class _Group:
    processing_time: ProcessingTime
    samples: list[str]
    remaining: set[int]
    priced: list[PricedGroup] = field(default_factory=list)


def price(
    samples: tuple[BillableSample, ...],
    customer_quotes: list[Quote],
    default_quotes: list[Quote],
    on: date,
) -> list[PricedGroup]:
    groups: dict[tuple, _Group] = {}
    for sample in samples:
        key = (sample.processing_time, tuple(sorted(sample.analysis_ids)))
        group = groups.setdefault(key, _Group(sample.processing_time, [], set(sample.analysis_ids)))
        group.samples.append(sample.name)
    for quotes in (customer_quotes, default_quotes):
        ordered = sorted((q for q in quotes if q.valid_on(on)), key=lambda q: q.widest_line, reverse=True)
        for group in groups.values():
            for quote in ordered:
                if not group.remaining:
                    break
                _apply(quote, group)
    unpriced = sorted({a for g in groups.values() for a in g.remaining})
    if unpriced:
        raise Unpriced(f"no valid quote prices analyses {unpriced}")
    return _merge([p for g in groups.values() for p in g.priced])


def _apply(quote: Quote, group: _Group) -> None:
    quantity = len(group.samples)
    for line in quote.line_items:
        wanted = set(line.analysis_ids)
        quote_price = line.price_for(group.processing_time)
        if not wanted <= group.remaining or quote_price is None:
            continue
        bulk = line.bulk_price_min > 0 and quantity >= line.bulk_price_min and quote_price.bulk_price is not None
        group.priced.append(
            PricedGroup(
                processing_time=group.processing_time,
                quantity=quantity,
                samples=tuple(group.samples),
                analysis_ids=frozenset(wanted),
                quote_price_id=quote_price.id,
                price=quote_price.bulk_price if bulk else quote_price.regular_price,
                payment_terms=quote.payment_terms,
            )
        )
        group.remaining -= wanted
        if not group.remaining:
            return


def _merge(groups: list[PricedGroup]) -> list[PricedGroup]:
    by_price: dict[int, list[PricedGroup]] = {}
    for g in groups:
        by_price.setdefault(g.quote_price_id, []).append(g)
    merged = []
    for same in by_price.values():
        first = same[0]
        merged.append(
            PricedGroup(
                processing_time=first.processing_time,
                quantity=sum(g.quantity for g in same),
                samples=tuple(s for g in same for s in g.samples),
                analysis_ids=first.analysis_ids,
                quote_price_id=first.quote_price_id,
                price=first.price,
                payment_terms=first.payment_terms,
            )
        )
    return merged
