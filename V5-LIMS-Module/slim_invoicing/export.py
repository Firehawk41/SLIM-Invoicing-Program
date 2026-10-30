"""The accounting import file, ported from V4's CSV writer: one row per line item, same columns.

A text cell a spreadsheet would run as a formula (a leading = + - @, or a tab or CR) is written as text
(slim-data D-118): the PO number and the requester come from the customer's request form. Numbers stay
numbers.
"""

import csv
from collections.abc import Iterable
from decimal import Decimal, InvalidOperation
from typing import TextIO

from slim_invoicing.domain.sales_order import SalesOrder

HEADERS = (
    "Customer",
    "Date",
    "Item",
    "Description",
    "Quantity",
    "Rate",
    "PO Number",
    "Name of Requester",
    "Email of Requester",
    "Service Date",
    "Submission Ref",
)


def write_csv(orders: Iterable[SalesOrder], out: TextIO) -> int:
    """Writes the orders to `out`; returns the number of rows written (headers aside)."""
    writer = _SafeWriter(csv.writer(out))
    writer.writerow(HEADERS)
    rows = 0
    for order in orders:
        for li in order.line_items:
            writer.writerow(
                (
                    order.customer_name,
                    order.date_received.isoformat(),
                    li.item,
                    li.description,
                    li.quantity,
                    str(li.rate),
                    order.po_number,
                    order.requester_name,
                    order.requester_email,
                    order.service_date.isoformat(),
                    order.tr_number,
                )
            )
            rows += 1
    return rows


class _SafeWriter:
    def __init__(self, writer) -> None:
        self._writer = writer

    def writerow(self, row) -> None:
        self._writer.writerow([as_text(cell) for cell in row])


def as_text(cell):
    """The cell, or "'" + the cell when a spreadsheet would run it as a formula."""
    if not isinstance(cell, str) or cell[:1] not in ("=", "+", "-", "@", "\t", "\r"):
        return cell
    try:
        Decimal(cell)
        return cell
    except InvalidOperation:
        return "'" + cell
