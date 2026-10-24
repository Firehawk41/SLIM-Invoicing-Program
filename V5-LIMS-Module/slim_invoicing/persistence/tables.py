"""Invoicing's own tables (D-83): quotes and sales orders. Everything else is slim-domain's.

Portable (SQLite + Postgres): enums as ints, money as Numeric, no foreign keys into other modules'
tables (customer and analysis IDs are slim-domain's, stored as plain integers).
"""

from sqlalchemy import (
    JSON,
    Boolean,
    Column,
    Date,
    DateTime,
    ForeignKey,
    Integer,
    MetaData,
    Numeric,
    String,
    Table,
    Text,
    UniqueConstraint,
    func,
)

metadata = MetaData()
MONEY = Numeric(12, 2)

quotes = Table(
    "inv_quotes",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("effective_date", Date, nullable=False),
    Column("expire_date", Date, nullable=False),
    Column("payment_method", Integer, nullable=False),
    Column("payment_terms", Integer, nullable=False),
    Column("is_default", Boolean, nullable=False),
    Column("created_by", String(100), nullable=False),
    Column("comments", Text, nullable=False, default=""),
    Column("created_at", DateTime(timezone=True), nullable=False, server_default=func.now()),
)

quote_customers = Table(
    "inv_quote_customers",
    metadata,
    Column("quote_id", Integer, ForeignKey("inv_quotes.id"), primary_key=True),
    Column("customer_id", Integer, primary_key=True),  # slim-domain's customer ID
)

quote_lines = Table(
    "inv_quote_lines",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("quote_id", Integer, ForeignKey("inv_quotes.id"), nullable=False, index=True),
    Column("position", Integer, nullable=False),
    Column("bulk_price_min", Integer, nullable=False),
)

quote_line_analyses = Table(
    "inv_quote_line_analyses",
    metadata,
    Column("line_id", Integer, ForeignKey("inv_quote_lines.id"), primary_key=True),
    Column("analysis_id", Integer, primary_key=True),  # slim-domain's analysis ID
)

quote_prices = Table(
    "inv_quote_prices",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("line_id", Integer, ForeignKey("inv_quote_lines.id"), nullable=False, index=True),
    Column("processing_time", Integer, nullable=False),
    Column("regular_price", MONEY, nullable=False),
    Column("bulk_price", MONEY),
    UniqueConstraint("line_id", "processing_time"),
)

sales_orders = Table(
    "inv_sales_orders",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("tr_number", String(20), nullable=False, unique=True),  # one sales order per TR
    Column("customer_id", Integer, nullable=False, index=True),
    Column("customer_name", String(200), nullable=False),
    Column("payment_method", Integer, nullable=False),
    Column("payment_terms", Integer, nullable=False),
    Column("po_number", String(100), nullable=False),
    Column("requester_name", String(200), nullable=False),
    Column("requester_email", String(254), nullable=False),
    Column("date_received", Date, nullable=False),
    Column("service_date", Date, nullable=False),
    Column("created_by", String(100), nullable=False),
    Column("created_at", DateTime(timezone=True), nullable=False, server_default=func.now()),
)

sales_order_lines = Table(
    "inv_sales_order_lines",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("sales_order_id", Integer, ForeignKey("inv_sales_orders.id"), nullable=False, index=True),
    Column("position", Integer, nullable=False),
    Column("item", String(300), nullable=False),
    Column("description", Text, nullable=False),
    Column("quantity", Integer, nullable=False),
    Column("rate", MONEY, nullable=False),
    Column("quote_price_id", Integer, nullable=False),
    Column("processing_time", Integer, nullable=False),
    Column("samples", JSON, nullable=False),  # sample names
)
