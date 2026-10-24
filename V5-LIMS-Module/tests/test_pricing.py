"""The pricing engine behaves as V4's did (D-83): V4's end-to-end case, then each rule."""

from decimal import Decimal

import pytest

from slim_invoicing.pricing import Unpriced, price
from tests.factories import ANIONS, ELEMENTS, FIVE, ND, PH, TODAY, line, quote, sample

BUNDLE = line(ELEMENTS, ANIONS, prices=[(ND, "85", "70"), (FIVE, "55", None)], bulk_min=10)
CUSTOMER = quote(BUNDLE, customers=(7,), terms=45)
DEFAULTS = quote(
    line(PH, prices=[(ND, "18", None)]),
    line(ELEMENTS, prices=[(ND, "60", None)]),
    line(ANIONS, prices=[(ND, "40", None)]),
    default=True,
)


def test_v4s_end_to_end_case():
    """V4's sample run: W-001 + W-002 merge on the bundle, W-003 pH, W-004 the bundle at five days."""
    samples = (
        sample(1, ELEMENTS, ANIONS),
        sample(2, ELEMENTS, ANIONS),
        sample(3, PH),
        sample(4, ELEMENTS, ANIONS, pt=FIVE),
    )
    groups = {(g.samples, g.price) for g in price(samples, [CUSTOMER], [DEFAULTS], TODAY)}
    assert groups == {(("W-001", "W-002"), Decimal("85")), (("W-003",), Decimal("18")), (("W-004",), Decimal("55"))}


def test_the_bundle_wins_over_its_parts_and_the_customer_over_defaults():
    [group] = price((sample(1, ELEMENTS, ANIONS),), [CUSTOMER], [DEFAULTS], TODAY)
    assert group.analysis_ids == {ELEMENTS, ANIONS} and group.payment_terms == 45


def test_what_the_customer_quote_misses_falls_to_the_defaults():
    groups = price((sample(1, ELEMENTS, ANIONS, PH),), [CUSTOMER], [DEFAULTS], TODAY)
    assert {(frozenset(g.analysis_ids), g.price) for g in groups} == {
        (frozenset({ELEMENTS, ANIONS}), Decimal("85")),
        (frozenset({PH}), Decimal("18")),
    }


def test_bulk_price_from_the_lines_minimum():
    samples = tuple(sample(i, ELEMENTS, ANIONS) for i in range(1, 11))
    [group] = price(samples, [CUSTOMER], [], TODAY)
    assert group.quantity == 10 and group.price == Decimal("70")
    [few] = price(samples[:9], [CUSTOMER], [], TODAY)
    assert few.price == Decimal("85")


def test_expired_quotes_and_missing_turnarounds_never_price():
    expired = quote(line(PH, prices=[(ND, "1", None)]), default=True, end=TODAY.replace(day=1))
    with pytest.raises(Unpriced, match=r"\[3\]"):
        price((sample(1, PH),), [], [expired], TODAY)
    with pytest.raises(Unpriced):
        price((sample(1, PH, pt=FIVE),), [], [DEFAULTS], TODAY)  # pH has no five-day price


# D-84: bulk on the merged quantity; the customer pays the lowest total.

ELEMENTS_10_PLUS = line(ELEMENTS, prices=[(ND, "85", "70")], bulk_min=10)
PH_ONLY = line(PH, prices=[(ND, "18", None)])


def test_bulk_counts_the_merged_line():
    """Jamie's example: 6 × 36 Elements and 6 × 36 Elements + pH are one line of 12 × 36 Elements."""
    samples = tuple(sample(i, ELEMENTS) for i in range(1, 7)) + tuple(sample(i, ELEMENTS, PH) for i in range(7, 13))
    groups = {
        frozenset(g.analysis_ids): (g.quantity, g.price)
        for g in price(samples, [], [quote(ELEMENTS_10_PLUS, PH_ONLY, default=True)], TODAY)
    }
    assert groups == {frozenset({ELEMENTS}): (12, Decimal("70")), frozenset({PH}): (6, Decimal("18"))}


def test_the_parts_win_when_their_bulk_price_makes_them_cheaper_than_the_bundle():
    """12 samples: the bundle at $85 (no bulk) costs $1,020; the parts at $70 + $10 cost $960."""
    bundle = line(ELEMENTS, PH, prices=[(ND, "85", None)])
    parts = quote(ELEMENTS_10_PLUS, line(PH, prices=[(ND, "12", "10")], bulk_min=10), default=True)
    samples = tuple(sample(i, ELEMENTS, PH) for i in range(1, 13))
    lines = price(samples, [], [quote(bundle, default=True), parts], TODAY)
    assert sum(g.quantity * g.price for g in lines) == Decimal("960") and len(lines) == 2


def test_the_bundle_wins_when_it_is_cheaper():
    bundle = line(ELEMENTS, PH, prices=[(ND, "75", None)])  # vs 85 + 18 for the parts
    [only] = price((sample(1, ELEMENTS, PH),), [], [quote(bundle, ELEMENTS_10_PLUS, PH_ONLY, default=True)], TODAY)
    assert only.analysis_ids == {ELEMENTS, PH} and only.price == Decimal("75")


def test_a_tie_goes_to_fewer_lines():
    bundle = line(ELEMENTS, PH, prices=[(ND, "103", None)])  # = 85 + 18
    [only] = price((sample(1, ELEMENTS, PH),), [], [quote(ELEMENTS_10_PLUS, PH_ONLY, bundle, default=True)], TODAY)
    assert only.analysis_ids == {ELEMENTS, PH}


def test_the_customers_own_lines_still_come_before_cheaper_defaults():
    """V4's priority stays: a default is used only for what the customer's quotes don't price."""
    cheap = quote(line(ELEMENTS, prices=[(ND, "1", None)]), default=True)
    [only] = price((sample(1, ELEMENTS),), [quote(ELEMENTS_10_PLUS, customers=(7,))], [cheap], TODAY)
    assert only.price == Decimal("85")


def test_too_many_combinations_fall_back_to_each_groups_cheapest_cover(monkeypatch):
    import slim_invoicing.pricing as pricing

    monkeypatch.setattr(pricing, "MAX_COMBINATIONS", 0)
    bundle = line(ELEMENTS, PH, prices=[(ND, "75", None)])
    [only] = price((sample(1, ELEMENTS, PH),), [], [quote(bundle, ELEMENTS_10_PLUS, PH_ONLY, default=True)], TODAY)
    assert only.price == Decimal("75")
