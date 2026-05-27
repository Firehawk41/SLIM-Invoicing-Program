from datetime import date
from decimal import Decimal
from typing import Optional

from sqlalchemy import Numeric, String, text
from sqlalchemy.orm import Mapped, Session, mapped_column

from infrastructure.database import Base
from slim.domain.quote.enums import PaymentType
from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_line_item import QuoteLineItem
from slim.domain.quote.quote_price import QuotePrice
from slim.domain.tr.enums import ProcessingTime, RequestType

# ---------------------------------------------------------------------------
# ORM rows — customer quote tables
# ---------------------------------------------------------------------------

class _QuoteRow(Base):
    __tablename__ = "quotes"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    effective_date: Mapped[date] = mapped_column("effective_date")
    expiry_date: Mapped[date] = mapped_column("expiry_date")
    payment_type: Mapped[int] = mapped_column("payment_type")
    payment_value: Mapped[int] = mapped_column("payment_value")
    valid: Mapped[bool] = mapped_column("valid")
    created_by: Mapped[str] = mapped_column(String)
    created_date: Mapped[date] = mapped_column("created_date")
    modified_date: Mapped[Optional[date]] = mapped_column("modified_date", nullable=True)
    comments: Mapped[Optional[str]] = mapped_column(String, nullable=True)


class _QuoteLineItemRow(Base):
    __tablename__ = "quote_line_items"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_id: Mapped[int] = mapped_column("quote_id")
    request_type_id: Mapped[int] = mapped_column("request_type_id")
    chemical_id: Mapped[Optional[int]] = mapped_column("chemical_id", nullable=True)
    bulk_min: Mapped[int] = mapped_column("bulk_min")
    valid: Mapped[bool] = mapped_column("valid")


class _QuotePriceRow(Base):
    __tablename__ = "quote_prices"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("quote_line_item_id")
    processing_time_id: Mapped[int] = mapped_column("processing_time_id")
    regular_price: Mapped[Decimal] = mapped_column(Numeric(10, 4, asdecimal=True))
    bulk_price: Mapped[Decimal] = mapped_column(Numeric(10, 4, asdecimal=True))
    valid: Mapped[bool] = mapped_column("valid")


class _QuoteAnalysisRow(Base):
    __tablename__ = "quote_analysis"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("quote_line_item_id")
    analysis_id: Mapped[int] = mapped_column("analysis_id")


class _QuoteLineItemElementRow(Base):
    __tablename__ = "quote_line_item_element"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("quote_line_item_id")
    element_id: Mapped[int] = mapped_column("element_id")


class _QuoteCustomerRow(Base):
    __tablename__ = "quote_customer"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_id: Mapped[int] = mapped_column("quote_id")
    customer_id: Mapped[int] = mapped_column("customer_id")


# ---------------------------------------------------------------------------
# ORM rows — default quote tables
# Python attribute names mirror their customer-quote equivalents so the
# cache helpers work uniformly across both table sets.
# ---------------------------------------------------------------------------

class _DefaultQuoteRow(Base):
    __tablename__ = "default_quotes"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    effective_date: Mapped[date] = mapped_column("effective_date")
    expiry_date: Mapped[date] = mapped_column("expiry_date")
    payment_type: Mapped[int] = mapped_column("payment_type")
    payment_value: Mapped[int] = mapped_column("payment_value")
    valid: Mapped[bool] = mapped_column("valid")
    created_by: Mapped[str] = mapped_column(String)
    created_date: Mapped[date] = mapped_column("created_date")
    modified_date: Mapped[Optional[date]] = mapped_column("modified_date", nullable=True)
    comments: Mapped[Optional[str]] = mapped_column(String, nullable=True)


class _DefaultQuoteLineItemRow(Base):
    __tablename__ = "default_quote_line_items"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_id: Mapped[int] = mapped_column("default_quote_id")  # same Python attr, different DB col
    request_type_id: Mapped[int] = mapped_column("request_type_id")
    chemical_id: Mapped[Optional[int]] = mapped_column("chemical_id", nullable=True)
    bulk_min: Mapped[int] = mapped_column("bulk_min")
    valid: Mapped[bool] = mapped_column("valid")


class _DefaultQuotePriceRow(Base):
    __tablename__ = "default_quote_prices"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("default_quote_line_item_id")
    processing_time_id: Mapped[int] = mapped_column("processing_time_id")
    regular_price: Mapped[Decimal] = mapped_column(Numeric(10, 4, asdecimal=True))
    bulk_price: Mapped[Decimal] = mapped_column(Numeric(10, 4, asdecimal=True))
    valid: Mapped[bool] = mapped_column("valid")


class _DefaultQuoteAnalysisRow(Base):
    __tablename__ = "default_quote_analysis"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("default_quote_line_item_id")
    analysis_id: Mapped[int] = mapped_column("analysis_id")


class _DefaultQuoteLineItemElementRow(Base):
    __tablename__ = "default_quote_line_item_element"
    id: Mapped[int] = mapped_column("ID", primary_key=True, autoincrement=True)
    quote_line_item_id: Mapped[int] = mapped_column("default_quote_line_item_id")
    element_id: Mapped[int] = mapped_column("element_id")


# ---------------------------------------------------------------------------
# Repository
# ---------------------------------------------------------------------------

class QuoteRepository:
    """Owns all SQL for the quote domain. Returns ORM rows for reads."""

    def __init__(self, session: Session) -> None:
        self._session = session

    # --- reads: customer quotes ---

    def select_all_quotes(self) -> list[_QuoteRow]:
        return self._session.query(_QuoteRow).order_by(_QuoteRow.id).all()

    def select_all_line_items(self) -> list[_QuoteLineItemRow]:
        return self._session.query(_QuoteLineItemRow).order_by(_QuoteLineItemRow.id).all()

    def select_all_prices(self) -> list[_QuotePriceRow]:
        return self._session.query(_QuotePriceRow).order_by(_QuotePriceRow.id).all()

    def select_all_analyses(self) -> list[_QuoteAnalysisRow]:
        return self._session.query(_QuoteAnalysisRow).order_by(_QuoteAnalysisRow.id).all()

    def select_all_elements(self) -> list[_QuoteLineItemElementRow]:
        return self._session.query(_QuoteLineItemElementRow).order_by(_QuoteLineItemElementRow.id).all()

    def select_all_customers(self) -> list[_QuoteCustomerRow]:
        return self._session.query(_QuoteCustomerRow).order_by(_QuoteCustomerRow.id).all()

    # --- reads: default quotes ---

    def select_all_default_quotes(self) -> list[_DefaultQuoteRow]:
        return self._session.query(_DefaultQuoteRow).order_by(_DefaultQuoteRow.id).all()

    def select_all_default_line_items(self) -> list[_DefaultQuoteLineItemRow]:
        return self._session.query(_DefaultQuoteLineItemRow).order_by(_DefaultQuoteLineItemRow.id).all()

    def select_all_default_prices(self) -> list[_DefaultQuotePriceRow]:
        return self._session.query(_DefaultQuotePriceRow).order_by(_DefaultQuotePriceRow.id).all()

    def select_all_default_analyses(self) -> list[_DefaultQuoteAnalysisRow]:
        return self._session.query(_DefaultQuoteAnalysisRow).order_by(_DefaultQuoteAnalysisRow.id).all()

    def select_all_default_elements(self) -> list[_DefaultQuoteLineItemElementRow]:
        return self._session.query(_DefaultQuoteLineItemElementRow).order_by(_DefaultQuoteLineItemElementRow.id).all()

    # --- writes: customer quotes ---

    def insert_quote(self, quote: Quote) -> int:
        row = _QuoteRow(
            effective_date=quote.effective_date,
            expiry_date=quote.expire_date,
            payment_type=int(quote.payment_type),
            payment_value=quote.payment_terms,
            valid=quote.is_valid,
            created_by=quote.created_by,
            created_date=quote.created_date,
            modified_date=quote.modified_date,
            comments=quote.comments,
        )
        self._session.add(row)
        self._session.flush()
        quote_id = row.id
        self._insert_line_items(quote.line_items, quote_id, is_default=False)
        for cid in quote.customer_ids:
            self._session.add(_QuoteCustomerRow(quote_id=quote_id, customer_id=cid))
        self._session.commit()
        return quote_id

    def update_quote(self, quote: Quote) -> None:
        self._session.execute(
            text(
                "UPDATE quotes SET "
                "effective_date=:ed, expiry_date=:xd, payment_type=:pt, "
                "payment_value=:pv, valid=:v, created_by=:cb, "
                "created_date=:cd, modified_date=:md, comments=:co "
                "WHERE ID=:qid"
            ),
            {
                "ed": quote.effective_date,
                "xd": quote.expire_date,
                "pt": int(quote.payment_type),
                "pv": quote.payment_terms,
                "v": 1 if quote.is_valid else 0,
                "cb": quote.created_by,
                "cd": quote.created_date,
                "md": quote.modified_date,
                "co": quote.comments,
                "qid": quote.id,
            },
        )
        self._session.commit()

    def delete_quote(self, quote_id: int) -> None:
        self._delete_line_items(quote_id, is_default=False)
        self._session.execute(
            text("DELETE FROM quote_customer WHERE quote_id=:qid"), {"qid": quote_id}
        )
        self._session.execute(
            text("DELETE FROM quotes WHERE ID=:qid"), {"qid": quote_id}
        )
        self._session.commit()

    # --- writes: default quotes ---

    def insert_default_quote(self, quote: Quote) -> int:
        row = _DefaultQuoteRow(
            effective_date=quote.effective_date,
            expiry_date=quote.expire_date,
            payment_type=int(quote.payment_type),
            payment_value=quote.payment_terms,
            valid=quote.is_valid,
            created_by=quote.created_by,
            created_date=quote.created_date,
            modified_date=quote.modified_date,
            comments=quote.comments,
        )
        self._session.add(row)
        self._session.flush()
        quote_id = row.id
        self._insert_line_items(quote.line_items, quote_id, is_default=True)
        self._session.commit()
        return quote_id

    def delete_default_quote(self, quote_id: int) -> None:
        self._delete_line_items(quote_id, is_default=True)
        self._session.execute(
            text("DELETE FROM default_quotes WHERE ID=:qid"), {"qid": quote_id}
        )
        self._session.commit()

    # --- private helpers ---

    def _insert_line_items(
        self, line_items: tuple[QuoteLineItem, ...], quote_id: int, is_default: bool
    ) -> None:
        for li in line_items:
            li_id = self._insert_line_item(li, quote_id, is_default)
            self._insert_prices(li.prices, li_id, is_default)
            self._insert_analyses(li.analysis_ids, li_id, is_default)
            self._insert_elements(li.element_ids, li_id, is_default)

    def _insert_line_item(self, li: QuoteLineItem, quote_id: int, is_default: bool) -> int:
        RowClass = _DefaultQuoteLineItemRow if is_default else _QuoteLineItemRow
        row = RowClass(
            quote_id=quote_id,
            request_type_id=int(li.request_type),
            chemical_id=li.chemical_id,
            bulk_min=li.bulk_price_min,
            valid=li.is_valid,
        )
        self._session.add(row)
        self._session.flush()
        return row.id

    def _insert_prices(
        self, prices: tuple[QuotePrice, ...], li_id: int, is_default: bool
    ) -> None:
        RowClass = _DefaultQuotePriceRow if is_default else _QuotePriceRow
        for p in prices:
            self._session.add(
                RowClass(
                    quote_line_item_id=li_id,
                    processing_time_id=int(p.processing_time),
                    regular_price=p.regular_price,
                    bulk_price=p.bulk_price,
                    valid=p.is_valid,
                )
            )

    def _insert_analyses(self, analysis_ids: tuple[int, ...], li_id: int, is_default: bool) -> None:
        RowClass = _DefaultQuoteAnalysisRow if is_default else _QuoteAnalysisRow
        for aid in analysis_ids:
            self._session.add(RowClass(quote_line_item_id=li_id, analysis_id=aid))

    def _insert_elements(self, element_ids: tuple[int, ...], li_id: int, is_default: bool) -> None:
        RowClass = _DefaultQuoteLineItemElementRow if is_default else _QuoteLineItemElementRow
        for eid in element_ids:
            self._session.add(RowClass(quote_line_item_id=li_id, element_id=eid))

    def _delete_line_items(self, quote_id: int, is_default: bool) -> None:
        if is_default:
            li_table, price_table = "default_quote_line_items", "default_quote_prices"
            analysis_table, element_table = "default_quote_analysis", "default_quote_line_item_element"
            fk_col, child_fk = "default_quote_id", "default_quote_line_item_id"
        else:
            li_table, price_table = "quote_line_items", "quote_prices"
            analysis_table, element_table = "quote_analysis", "quote_line_item_element"
            fk_col, child_fk = "quote_id", "quote_line_item_id"

        sub = f"(SELECT ID FROM {li_table} WHERE {fk_col}=:qid)"
        for child in (price_table, analysis_table, element_table):
            self._session.execute(
                text(f"DELETE FROM {child} WHERE {child_fk} IN {sub}"), {"qid": quote_id}
            )
        self._session.execute(
            text(f"DELETE FROM {li_table} WHERE {fk_col}=:qid"), {"qid": quote_id}
        )
