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
