"""Quotes and sales orders (D-83).

Revision ID: 0001
Revises:
Create Date: 2026-10-24
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0001"
down_revision: str | None = None
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None

MONEY = sa.Numeric(12, 2)


def upgrade() -> None:
    op.create_table(
        "inv_quotes",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("effective_date", sa.Date, nullable=False),
        sa.Column("expire_date", sa.Date, nullable=False),
        sa.Column("payment_method", sa.Integer, nullable=False),
        sa.Column("payment_terms", sa.Integer, nullable=False),
        sa.Column("is_default", sa.Boolean, nullable=False),
        sa.Column("created_by", sa.String(100), nullable=False),
        sa.Column("comments", sa.Text, nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.func.now()),
    )
    op.create_table(
        "inv_quote_customers",
        sa.Column("quote_id", sa.Integer, sa.ForeignKey("inv_quotes.id"), primary_key=True),
        sa.Column("customer_id", sa.Integer, primary_key=True),
    )
    op.create_table(
        "inv_quote_lines",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("quote_id", sa.Integer, sa.ForeignKey("inv_quotes.id"), nullable=False, index=True),
        sa.Column("position", sa.Integer, nullable=False),
        sa.Column("bulk_price_min", sa.Integer, nullable=False),
    )
    op.create_table(
        "inv_quote_line_analyses",
        sa.Column("line_id", sa.Integer, sa.ForeignKey("inv_quote_lines.id"), primary_key=True),
        sa.Column("analysis_id", sa.Integer, primary_key=True),
    )
    op.create_table(
        "inv_quote_prices",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("line_id", sa.Integer, sa.ForeignKey("inv_quote_lines.id"), nullable=False, index=True),
        sa.Column("processing_time", sa.Integer, nullable=False),
        sa.Column("regular_price", MONEY, nullable=False),
        sa.Column("bulk_price", MONEY),
        sa.UniqueConstraint("line_id", "processing_time"),
    )
    op.create_table(
        "inv_sales_orders",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("tr_number", sa.String(20), nullable=False, unique=True),
        sa.Column("customer_id", sa.Integer, nullable=False, index=True),
        sa.Column("customer_name", sa.String(200), nullable=False),
        sa.Column("payment_method", sa.Integer, nullable=False),
        sa.Column("payment_terms", sa.Integer, nullable=False),
        sa.Column("po_number", sa.String(100), nullable=False),
        sa.Column("requester_name", sa.String(200), nullable=False),
        sa.Column("requester_email", sa.String(254), nullable=False),
        sa.Column("date_received", sa.Date, nullable=False),
        sa.Column("service_date", sa.Date, nullable=False),
        sa.Column("created_by", sa.String(100), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.func.now()),
    )
    op.create_table(
        "inv_sales_order_lines",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("sales_order_id", sa.Integer, sa.ForeignKey("inv_sales_orders.id"), nullable=False, index=True),
        sa.Column("position", sa.Integer, nullable=False),
        sa.Column("item", sa.String(300), nullable=False),
        sa.Column("description", sa.Text, nullable=False),
        sa.Column("quantity", sa.Integer, nullable=False),
        sa.Column("rate", MONEY, nullable=False),
        sa.Column("quote_price_id", sa.Integer, nullable=False),
        sa.Column("processing_time", sa.Integer, nullable=False),
        sa.Column("samples", sa.JSON, nullable=False),
    )


def downgrade() -> None:
    for table in (
        "inv_sales_order_lines",
        "inv_sales_orders",
        "inv_quote_prices",
        "inv_quote_line_analyses",
        "inv_quote_lines",
        "inv_quote_customers",
        "inv_quotes",
    ):
        op.drop_table(table)
