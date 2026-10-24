"""Quotes and sales orders in and out of the database. Repositories flush; the host commits."""

from decimal import Decimal

from sqlalchemy import insert, select
from sqlalchemy.orm import Session

from slim_invoicing.domain.enums import PaymentMethod, ProcessingTime
from slim_invoicing.domain.quote import Quote, QuoteLineItem, QuotePrice
from slim_invoicing.domain.sales_order import SalesOrder, SalesOrderLineItem
from slim_invoicing.persistence import tables as t

# --------------------------------------------------------------------- quotes


def save_quote(session: Session, quote: Quote) -> Quote:
    """Stores a new quote and returns it with its IDs. Quotes are not edited: retire one by expiring it."""
    quote_id = session.execute(
        insert(t.quotes).values(
            effective_date=quote.effective_date,
            expire_date=quote.expire_date,
            payment_method=int(quote.payment_method),
            payment_terms=quote.payment_terms,
            is_default=quote.is_default,
            created_by=quote.created_by,
            comments=quote.comments,
        )
    ).inserted_primary_key[0]
    for customer_id in quote.customer_ids:
        session.execute(insert(t.quote_customers).values(quote_id=quote_id, customer_id=customer_id))
    for position, line in enumerate(quote.line_items, 1):
        line_id = session.execute(
            insert(t.quote_lines).values(quote_id=quote_id, position=position, bulk_price_min=line.bulk_price_min)
        ).inserted_primary_key[0]
        for analysis_id in line.analysis_ids:
            session.execute(insert(t.quote_line_analyses).values(line_id=line_id, analysis_id=analysis_id))
        for p in line.prices:
            session.execute(
                insert(t.quote_prices).values(
                    line_id=line_id,
                    processing_time=int(p.processing_time),
                    regular_price=p.regular_price,
                    bulk_price=p.bulk_price,
                )
            )
    session.flush()
    return _load_quotes(session, [quote_id])[0]


def quotes_for_customer(session: Session, customer_id: int) -> list[Quote]:
    ids = session.scalars(select(t.quote_customers.c.quote_id).where(t.quote_customers.c.customer_id == customer_id))
    return _load_quotes(session, list(ids))


def default_quotes(session: Session) -> list[Quote]:
    return _load_quotes(session, list(session.scalars(select(t.quotes.c.id).where(t.quotes.c.is_default))))


def _load_quotes(session: Session, ids: list[int]) -> list[Quote]:
    if not ids:
        return []
    customers: dict[int, list[int]] = {}
    for row in session.execute(select(t.quote_customers).where(t.quote_customers.c.quote_id.in_(ids))):
        customers.setdefault(row.quote_id, []).append(row.customer_id)
    lines = session.execute(
        select(t.quote_lines).where(t.quote_lines.c.quote_id.in_(ids)).order_by(t.quote_lines.c.position)
    ).all()
    line_ids = [line.id for line in lines]
    analyses: dict[int, list[int]] = {}
    for row in session.execute(
        select(t.quote_line_analyses)
        .where(t.quote_line_analyses.c.line_id.in_(line_ids))
        .order_by(t.quote_line_analyses.c.analysis_id)
    ):
        analyses.setdefault(row.line_id, []).append(row.analysis_id)
    prices: dict[int, list[QuotePrice]] = {}
    for row in session.execute(
        select(t.quote_prices).where(t.quote_prices.c.line_id.in_(line_ids)).order_by(t.quote_prices.c.id)
    ):
        prices.setdefault(row.line_id, []).append(
            QuotePrice(
                id=row.id,
                processing_time=ProcessingTime(row.processing_time),
                regular_price=Decimal(row.regular_price),
                bulk_price=None if row.bulk_price is None else Decimal(row.bulk_price),
            )
        )
    by_quote: dict[int, list[QuoteLineItem]] = {}
    for line in lines:
        by_quote.setdefault(line.quote_id, []).append(
            QuoteLineItem(
                id=line.id,
                analysis_ids=tuple(analyses.get(line.id, ())),
                bulk_price_min=line.bulk_price_min,
                prices=tuple(prices.get(line.id, ())),
            )
        )
    rows = session.execute(select(t.quotes).where(t.quotes.c.id.in_(ids)).order_by(t.quotes.c.id))
    return [
        Quote(
            id=row.id,
            effective_date=row.effective_date,
            expire_date=row.expire_date,
            payment_method=PaymentMethod(row.payment_method),
            payment_terms=row.payment_terms,
            is_default=row.is_default,
            customer_ids=tuple(sorted(customers.get(row.id, ()))),
            line_items=tuple(by_quote[row.id]),
            created_by=row.created_by,
            comments=row.comments,
        )
        for row in rows
    ]


# --------------------------------------------------------------------- sales orders


def save_sales_order(session: Session, order: SalesOrder) -> SalesOrder:
    """Stores a new sales order. The unique TR number refuses a second one for the same TR."""
    order_id = session.execute(
        insert(t.sales_orders).values(
            tr_number=order.tr_number,
            customer_id=order.customer_id,
            customer_name=order.customer_name,
            payment_method=int(order.payment_method),
            payment_terms=order.payment_terms,
            po_number=order.po_number,
            requester_name=order.requester_name,
            requester_email=order.requester_email,
            date_received=order.date_received,
            service_date=order.service_date,
            created_by=order.created_by,
        )
    ).inserted_primary_key[0]
    for position, li in enumerate(order.line_items, 1):
        session.execute(
            insert(t.sales_order_lines).values(
                sales_order_id=order_id,
                position=position,
                item=li.item,
                description=li.description,
                quantity=li.quantity,
                rate=li.rate,
                quote_price_id=li.quote_price_id,
                processing_time=int(li.processing_time),
                samples=list(li.samples),
            )
        )
    session.flush()
    return sales_order_for(session, order.tr_number)


def sales_order_for(session: Session, tr_number: str) -> SalesOrder | None:
    row = session.execute(select(t.sales_orders).where(t.sales_orders.c.tr_number == tr_number)).first()
    return None if row is None else _order(session, row)


def sales_orders(session: Session) -> list[SalesOrder]:
    """Every sales order, newest first."""
    rows = session.execute(select(t.sales_orders).order_by(t.sales_orders.c.id.desc())).all()
    return [_order(session, row) for row in rows]


def _order(session: Session, row) -> SalesOrder:
    lines = session.execute(
        select(t.sales_order_lines)
        .where(t.sales_order_lines.c.sales_order_id == row.id)
        .order_by(t.sales_order_lines.c.position)
    )
    return SalesOrder(
        id=row.id,
        tr_number=row.tr_number,
        customer_id=row.customer_id,
        customer_name=row.customer_name,
        payment_method=PaymentMethod(row.payment_method),
        payment_terms=row.payment_terms,
        po_number=row.po_number,
        requester_name=row.requester_name,
        requester_email=row.requester_email,
        date_received=row.date_received,
        service_date=row.service_date,
        created_by=row.created_by,
        created_at=row.created_at,
        line_items=tuple(
            SalesOrderLineItem(
                item=li.item,
                description=li.description,
                quantity=li.quantity,
                rate=Decimal(li.rate),
                quote_price_id=li.quote_price_id,
                processing_time=ProcessingTime(li.processing_time),
                samples=tuple(li.samples),
            )
            for li in lines
        ),
    )
