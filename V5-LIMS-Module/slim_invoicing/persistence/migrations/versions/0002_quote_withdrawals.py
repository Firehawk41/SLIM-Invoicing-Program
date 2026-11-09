"""Quote withdrawals (D-139): a quote withdrawn before its expiry, as a row added.

Revision ID: 0002
Revises: 0001
Create Date: 2026-11-09
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0002"
down_revision: str | None = "0001"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "inv_quote_withdrawals",
        sa.Column("quote_id", sa.Integer, sa.ForeignKey("inv_quotes.id"), primary_key=True),
        sa.Column("from_date", sa.Date, nullable=False),
        sa.Column("withdrawn_by", sa.String(100), nullable=False),
        sa.Column("reason", sa.Text, nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.func.now()),
    )


def downgrade() -> None:
    op.drop_table("inv_quote_withdrawals")
