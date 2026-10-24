"""Bill what was reported, once (D-83)."""

from tests.factories import ANIONS, ELEMENTS, PH, request, sample


def test_only_reported_analyses_are_billed():
    req = request(sample(1, ELEMENTS, ANIONS), sample(2, PH), reported={(1, ELEMENTS)})
    [billed] = req.billable_samples()
    assert billed.position == 1 and billed.analysis_ids == (ELEMENTS,)  # sample 2 was never reported


def test_an_amended_report_counts_once():
    """reported is a set: an analysis on the original and the amended issue is one (position, analysis)."""
    req = request(sample(1, ELEMENTS), reported=[(1, ELEMENTS), (1, ELEMENTS)])
    assert [s.analysis_ids for s in req.billable_samples()] == [(ELEMENTS,)]
