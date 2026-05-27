from datetime import date
from decimal import Decimal

from sqlalchemy import Numeric, String, text
from sqlalchemy.orm import Mapped, Session, mapped_column

from infrastructure.database import Base
from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem
from slim.domain.tr.enums import ProcessingTime


# ---------------------------------------------------------------------------
# ORM rows
# ---------------------------------------------------------------------------


class _SalesOrderRow(Base):
    __tablename__ = "sales_orders"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    customer_name: Mapped[str] = mapped_column(String)
    customer_id: Mapped[int] = mapped_column("customer_id")
    submission_id: Mapped[int] = mapped_column("submission_id")
    submission_ref: Mapped[str] = mapped_column(String)
    date_received: Mapped[date] = mapped_column("date_received")
    po_number: Mapped[str] = mapped_column(String)
    payment_type: Mapped[int] = mapped_column("payment_type")
    payment_terms: Mapped[int] = mapped_column("payment_terms")
    requester_name: Mapped[str] = mapped_column(String)
    requester_email: Mapped[str] = mapped_column(String)
    service_date: Mapped[date] = mapped_column("service_date")


class _SalesOrderLineItemRow(Base):
    __tablename__ = "sales_order_items"
    id: Mapped[int] = mapped_column("sales_order_item_id", primary_key=True, autoincrement=True)
    sales_order_id: Mapped[int] = mapped_column("sales_order_id")
    item: Mapped[str] = mapped_column(String)
    description: Mapped[str] = mapped_column(String)
    quantity: Mapped[int] = mapped_column("quantity")
    rate: Mapped[Decimal] = mapped_column(Numeric(10, 4, asdecimal=True))
    quote_price_id: Mapped[int] = mapped_column("quote_price_id")
    processing_time_id: Mapped[int] = mapped_column("processing_time_id")


# ---------------------------------------------------------------------------
# Line item repository (internal — not exposed via __init__.py)
# ---------------------------------------------------------------------------


class _SalesOrderLineItemRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def select_by_sales_order_id(
        self, sales_order_id: int
    ) -> list[_SalesOrderLineItemRow]:
        return (
            self._session.query(_SalesOrderLineItemRow)
            .filter(_SalesOrderLineItemRow.sales_order_id == sales_order_id)
            .order_by(_SalesOrderLineItemRow.id)
            .all()
        )

    def insert(self, li: SalesOrderLineItem, sales_order_id: int) -> int:
        row = _SalesOrderLineItemRow(
            sales_order_id=sales_order_id,
            item=li.item,
            description=li.description,
            quantity=li.quantity,
            rate=li.rate,
            quote_price_id=li.quote_price_id,
            processing_time_id=int(li.processing_time),
        )
        self._session.add(row)
        self._session.flush()
        return row.id

    def delete_by_sales_order_id(self, sales_order_id: int) -> None:
        self._session.execute(
            text("DELETE FROM sales_order_items WHERE sales_order_id=:sid"),
            {"sid": sales_order_id},
        )

    def delete(self, line_item_id: int) -> None:
        self._session.execute(
            text("DELETE FROM sales_order_items WHERE sales_order_item_id=:lid"),
            {"lid": line_item_id},
        )


# ---------------------------------------------------------------------------
# Sales order repository
# ---------------------------------------------------------------------------


class SalesOrderRepository:
    """Owns all SQL for sales_orders and sales_order_items. Manages transactions."""

    def __init__(self, session: Session) -> None:
        self._session = session
        self._li_repo = _SalesOrderLineItemRepository(session)

    def insert(self, so: SalesOrder) -> SalesOrder:
        """Persist header + line items, commit, return entity with DB-assigned IDs."""
        row = _SalesOrderRow(
            customer_name=so.customer_name,
            customer_id=so.customer_id,
            submission_id=so.submission_id,
            submission_ref=so.submission_ref,
            date_received=so.date_received,
            po_number=so.po_number,
            payment_type=int(so.payment_type),
            payment_terms=so.payment_terms,
            requester_name=so.requester_name,
            requester_email=so.requester_email,
            service_date=so.service_date,
        )
        self._session.add(row)
        self._session.flush()
        new_id = row.id

        sealed_items = []
        for li in so.line_items:
            li_id = self._li_repo.insert(li, new_id)
            sealed_items.append(li.model_copy(update={"id": li_id}))

        self._session.commit()
        return so.model_copy(update={"id": new_id, "line_items": tuple(sealed_items)})

    def update(self, so: SalesOrder) -> None:
        """Update header fields only. Line items are not touched."""
        self._session.execute(
            text(
                "UPDATE sales_orders SET "
                "customer_name=:cn, customer_id=:cid, submission_id=:sid, "
                "submission_ref=:sref, date_received=:dr, po_number=:po, "
                "payment_type=:pt, payment_terms=:pterm, "
                "requester_name=:rn, requester_email=:re, service_date=:sd "
                "WHERE ID=:id"
            ),
            {
                "cn": so.customer_name,
                "cid": so.customer_id,
                "sid": so.submission_id,
                "sref": so.submission_ref,
                "dr": so.date_received,
                "po": so.po_number,
                "pt": int(so.payment_type),
                "pterm": so.payment_terms,
                "rn": so.requester_name,
                "re": so.requester_email,
                "sd": so.service_date,
                "id": so.id,
            },
        )
        self._session.commit()

    def delete(self, sales_order_id: int) -> None:
        """Delete line items then header, commit."""
        self._li_repo.delete_by_sales_order_id(sales_order_id)
        self._session.execute(
            text("DELETE FROM sales_orders WHERE ID=:id"), {"id": sales_order_id}
        )
        self._session.commit()

    def load(self, sales_order_id: int) -> SalesOrder | None:
        row = (
            self._session.query(_SalesOrderRow)
            .filter(_SalesOrderRow.id == sales_order_id)
            .first()
        )
        if row is None:
            return None
        li_rows = self._li_repo.select_by_sales_order_id(sales_order_id)
        return self._build(row, li_rows)

    def load_by_date_range(self, start: date, end: date) -> list[SalesOrder]:
        rows = (
            self._session.query(_SalesOrderRow)
            .filter(
                _SalesOrderRow.service_date >= start,
                _SalesOrderRow.service_date <= end,
            )
            .order_by(_SalesOrderRow.service_date)
            .all()
        )
        return [self.load(row.id) for row in rows]  # type: ignore[misc]

    # --- private helpers ---

    def _build(
        self,
        row: _SalesOrderRow,
        li_rows: list[_SalesOrderLineItemRow],
    ) -> SalesOrder:
        line_items = tuple(
            SalesOrderLineItem(
                id=lr.id,
                item=lr.item,
                description=lr.description,
                quantity=lr.quantity,
                rate=Decimal(str(lr.rate)),
                quote_price_id=lr.quote_price_id,
                processing_time=ProcessingTime(lr.processing_time_id),
            )
            for lr in li_rows
        )
        return SalesOrder(
            id=row.id,
            customer_name=row.customer_name,
            customer_id=row.customer_id,
            submission_id=row.submission_id,
            submission_ref=row.submission_ref,
            date_received=row.date_received,
            po_number=row.po_number,
            payment_type=PaymentType(row.payment_type),
            payment_terms=row.payment_terms,
            requester_name=row.requester_name,
            requester_email=row.requester_email,
            service_date=row.service_date,
            line_items=line_items,
        )
