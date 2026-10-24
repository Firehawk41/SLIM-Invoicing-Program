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
