"""Invoice one TR (D-83): what was reported, priced from the quotes, as one sales order.

The host calls `invoice` in its own transaction, and moves the TR to invoiced in the same one, so a
sales order and the TR's status never disagree. Invoicing flushes; the host commits.
"""

from datetime import date

from sqlalchemy.orm import Session

from slim_invoicing import pricing
from slim_invoicing.domain.billable import BillableRequest
from slim_invoicing.domain.sales_order import SalesOrder, SalesOrderLineItem
from slim_invoicing.persistence import repositories as repo

DEFAULT_PAYMENT_TERMS = 30  # days, when no quote sets its own (V4's default)


class InvoiceRefused(ValueError):
    """The TR can't be invoiced as it stands; the message says why."""


def invoice(session: Session, request: BillableRequest, created_by: str, on: date | None = None) -> SalesOrder:
    on = on or date.today()
    if not request.complete:
        raise InvoiceRefused(f"{request.tr_number}: only a complete report is invoiced")
    if repo.sales_order_for(session, request.tr_number) is not None:
        raise InvoiceRefused(f"{request.tr_number} is already invoiced")
    samples = request.billable_samples()
    if not samples:
        raise InvoiceRefused(f"{request.tr_number}: nothing ordered was reported")
    try:
        groups = pricing.price(
            samples, repo.quotes_for_customer(session, request.customer_id), repo.default_quotes(session), on
        )
    except pricing.Unpriced as exc:
        raise InvoiceRefused(f"{request.tr_number}: {exc}") from exc
    terms = next((g.payment_terms for g in groups if g.payment_terms), DEFAULT_PAYMENT_TERMS)
    order = SalesOrder(
        tr_number=request.tr_number,
        customer_id=request.customer_id,
        customer_name=request.customer_name,
        payment_method=request.payment_method,
        payment_terms=terms,
        po_number=request.po_number,
        requester_name=request.requester_name,
        requester_email=request.requester_email,
        date_received=request.date_received,
        service_date=request.service_date,
        created_by=created_by,
        line_items=tuple(_line(g, request.analysis_names) for g in groups),
    )
    return repo.save_sales_order(session, order)


def _line(group: pricing.PricedGroup, names: dict[int, str]) -> SalesOrderLineItem:
    analyses = " + ".join(names.get(a, f"Analysis {a}") for a in sorted(group.analysis_ids))
    item = f"{analyses} - {group.processing_time.label}"
    return SalesOrderLineItem(
        item=item,
        description=f"{item}: {', '.join(group.samples)}",
        quantity=group.quantity,
        rate=group.price,
        quote_price_id=group.quote_price_id,
        processing_time=group.processing_time,
        samples=group.samples,
    )
