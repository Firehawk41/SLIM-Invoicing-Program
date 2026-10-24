"""Run invoicing's migrations on the host's connection (D-83).

Invoicing shares one database with the host and the other modules, so it keeps its own Alembic
version table and only ever touches its own `inv_*` tables. It never opens a connection itself.
"""

from pathlib import Path

from alembic import command
from alembic.config import Config
from sqlalchemy import Connection

VERSION_TABLE = "alembic_version_slim_invoicing"
_HERE = Path(__file__).resolve().parent


def alembic_config(connection: Connection) -> Config:
    config = Config(str(_HERE / "alembic.ini"))
    config.set_main_option("script_location", str(_HERE / "migrations"))
    config.attributes["connection"] = connection
    return config


def upgrade(connection: Connection, revision: str = "head") -> None:
    command.upgrade(alembic_config(connection), revision)


def downgrade(connection: Connection, revision: str = "base") -> None:
    command.downgrade(alembic_config(connection), revision)
