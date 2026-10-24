"""The accounting import file, ported from V4's CSV writer: one row per line item, same columns."""

import csv
from collections.abc import Iterable
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
    writer = csv.writer(out)
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
