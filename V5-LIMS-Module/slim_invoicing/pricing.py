"""The pricing engine: the cheapest correct invoice for what was reported (D-83, D-84).

Samples are grouped by turnaround + analysis set. Each group's analyses must be covered exactly by
quote lines that have a price for its turnaround: the customer's own quotes first, and the default
quotes only for analyses none of the customer's lines price (V4's priority). A group may be covered
by a bundle ("36 Elements + 7 Anions") or by its parts.

V5 changes two things in V4's greedy rule (Jamie, D-84):
  * the bulk price is decided on the merged quantity: every group priced on the same quote price
    becomes one line, and that line's total quantity is what meets the line's bulk minimum;
  * of all the ways to cover the groups, the customer pays the lowest total, so a bundle is used
    when it is cheaper than its parts and the parts when they are. Ties go to fewer lines.
A group left with an analysis no line prices is an error, never a zero-price line.
"""

from collections.abc import Iterator
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from itertools import product

from slim_invoicing.domain.billable import BillableSample
from slim_invoicing.domain.enums import ProcessingTime
from slim_invoicing.domain.quote import Quote

MAX_COMBINATIONS = 10_000  # beyond this, each group takes its own cheapest cover (a TR never gets close)


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


@dataclass(frozen=True)
class _Option:
    """One quote line at one turnaround: a way to price some of a group's analyses."""

    quote_price_id: int
    analysis_ids: frozenset[int]
    regular: Decimal
    bulk: Decimal | None
    bulk_min: int
    payment_terms: int

    def rate(self, quantity: int) -> Decimal:
        bulk = self.bulk is not None and self.bulk_min > 0 and quantity >= self.bulk_min
        return self.bulk if bulk else self.regular


@dataclass(frozen=True)
class _Group:
    processing_time: ProcessingTime
    analysis_ids: frozenset[int]
    samples: tuple[str, ...]


def price(
    samples: tuple[BillableSample, ...],
    customer_quotes: list[Quote],
    default_quotes: list[Quote],
    on: date,
) -> list[PricedGroup]:
    groups = _groups(samples)
    mine = [q for q in customer_quotes if q.valid_on(on)]
    defaults = [q for q in default_quotes if q.valid_on(on)]
    covers = [list(_covers(g.analysis_ids, _options(g, mine, defaults))) for g in groups]
    unpriced = sorted({a for g, c in zip(groups, covers, strict=True) if not c for a in g.analysis_ids})
    if unpriced:
        raise Unpriced(f"no valid quote prices analyses {unpriced}")
    combinations = 1
    for c in covers:
        combinations *= len(c)
    if combinations > MAX_COMBINATIONS:
        choices = [min(c, key=lambda cover, g=g: _cost([(g, cover)])) for g, c in zip(groups, covers, strict=True)]
        return _lines(list(zip(groups, choices, strict=True)))
    best = min(
        (list(zip(groups, choice, strict=True)) for choice in product(*covers)),
        key=_cost,
    )
    return _lines(best)


def _groups(samples: tuple[BillableSample, ...]) -> list[_Group]:
    names: dict[tuple, list[str]] = {}
    for s in samples:
        names.setdefault((s.processing_time, frozenset(s.analysis_ids)), []).append(s.name)
    return [_Group(pt, analyses, tuple(n)) for (pt, analyses), n in names.items()]


def _options(group: _Group, mine: list[Quote], defaults: list[Quote]) -> list[_Option]:
    def from_quotes(quotes: list[Quote], allowed: frozenset[int]) -> list[_Option]:
        out = []
        for q in quotes:
            for line in q.line_items:
                wanted = frozenset(line.analysis_ids)
                p = line.price_for(group.processing_time)
                if p is not None and wanted <= allowed:
                    out.append(
                        _Option(p.id, wanted, p.regular_price, p.bulk_price, line.bulk_price_min, q.payment_terms)
                    )
        return out

    own = from_quotes(mine, group.analysis_ids)
    covered = frozenset().union(*(o.analysis_ids for o in own))
    return own + from_quotes(defaults, group.analysis_ids - covered)


def _covers(remaining: frozenset[int], options: list[_Option]) -> Iterator[tuple[_Option, ...]]:
    """Every way to price `remaining` exactly once with the options (an exact cover)."""
    if not remaining:
        yield ()
        return
    first = min(remaining)
    for option in options:
        if first in option.analysis_ids and option.analysis_ids <= remaining:
            for rest in _covers(remaining - option.analysis_ids, options):
                yield (option, *rest)


def _merged(choice: list[tuple[_Group, tuple[_Option, ...]]]) -> dict[int, tuple[_Option, list[_Group]]]:
    lines: dict[int, tuple[_Option, list[_Group]]] = {}
    for group, cover in choice:
        for option in cover:
            lines.setdefault(option.quote_price_id, (option, []))[1].append(group)
    return lines


def _cost(choice: list[tuple[_Group, tuple[_Option, ...]]]) -> tuple:
    lines = _merged(choice)
    total = Decimal(0)
    for option, groups in lines.values():
        quantity = sum(len(g.samples) for g in groups)
        total += quantity * option.rate(quantity)
    return total, len(lines), sorted(lines)  # cheapest; then fewest lines; then a stable order


def _lines(choice: list[tuple[_Group, tuple[_Option, ...]]]) -> list[PricedGroup]:
    out = []
    for option, groups in _merged(choice).values():
        quantity = sum(len(g.samples) for g in groups)
        out.append(
            PricedGroup(
                processing_time=groups[0].processing_time,
                quantity=quantity,
                samples=tuple(s for g in groups for s in g.samples),
                analysis_ids=option.analysis_ids,
                quote_price_id=option.quote_price_id,
                price=option.rate(quantity),
                payment_terms=option.payment_terms,
            )
        )
    return out
