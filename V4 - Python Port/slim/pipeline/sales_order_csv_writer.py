import csv
from pathlib import Path
from types import TracebackType

from slim.domain.sales_order.sales_order import SalesOrder


class SalesOrderCsvWriter:
    """Context manager. Writes one CSV row per line item per sales order."""

    _HEADERS = [
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
    ]

    def __init__(self, path: str | Path) -> None:
        self._path = Path(path)
        self._file = None
        self._writer = None

    def __enter__(self) -> "SalesOrderCsvWriter":
        self._file = open(self._path, "w", newline="", encoding="utf-8")
        self._writer = csv.writer(self._file)
        self._writer.writerow(self._HEADERS)
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        if self._file:
            self._file.close()
            self._file = None
            self._writer = None

    def write(self, so: SalesOrder) -> None:
        assert self._writer is not None, "Must be used as a context manager."
        for li in so.line_items:
            self._writer.writerow(
                [
                    so.customer_name,
                    so.date_received.strftime("%Y-%m-%d"),
                    li.item,
                    li.description,
                    li.quantity,
                    str(li.rate),
                    so.po_number,
                    so.requester_name,
                    so.requester_email,
                    str(so.service_date),
                    so.submission_ref,
                ]
            )
