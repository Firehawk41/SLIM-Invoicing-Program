"""The accounting CSV keeps V4's columns (D-83)."""

import csv
import io

from slim_invoicing import invoice, save_quote, write_csv
from slim_invoicing.export import HEADERS
from tests.factories import ANIONS, ELEMENTS, PH, TODAY, request, sample
from tests.test_pricing import CUSTOMER, DEFAULTS


def test_one_row_per_line_item(session):
    save_quote(session, CUSTOMER)
    save_quote(session, DEFAULTS)
    order = invoice(session, request(sample(1, ELEMENTS, ANIONS), sample(2, PH)), "SUE", on=TODAY)
    out = io.StringIO()
    assert write_csv([order], out) == 2
    rows = list(csv.reader(io.StringIO(out.getvalue())))
    assert tuple(rows[0]) == HEADERS
    assert rows[1] == [
        "Acme Env Labs",
        "2026-10-20",
        "36 Elements + 7 Anions - Next Day",
        "36 Elements + 7 Anions - Next Day: W-001",
        "1",
        "85.00",
        "PO-1",
        "Jane Roe",
        "jane@acme.example",
        "2026-10-23",
        "TR00042",
    ]


def test_a_formula_from_the_request_form_is_written_as_text(session):  # slim-data D-118
    from slim_invoicing.export import as_text

    save_quote(session, CUSTOMER)
    save_quote(session, DEFAULTS)
    forged = request(sample(1, ELEMENTS)).model_copy(
        update={"requester_name": '=HYPERLINK("http://x","x")', "po_number": "@SUM(1)"}
    )
    out = io.StringIO()
    write_csv([invoice(session, forged, "SUE", on=TODAY)], out)
    row = list(csv.reader(io.StringIO(out.getvalue())))[1]
    assert row[HEADERS.index("Name of Requester")] == '\'=HYPERLINK("http://x","x")'
    assert row[HEADERS.index("PO Number")] == "'@SUM(1)"
    assert as_text("-12.50") == "-12.50" and as_text("-x") == "'-x" and as_text(3) == 3
