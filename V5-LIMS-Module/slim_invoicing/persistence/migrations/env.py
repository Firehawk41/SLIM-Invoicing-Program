"""Alembic environment: migrates only invoicing's own tables, on the host's connection.

Its own version table keeps its history apart from the host's and the other modules';
`include_object` keeps autogenerate away from tables it doesn't own.
"""

from alembic import context

from slim_invoicing.persistence.migrate import VERSION_TABLE
from slim_invoicing.persistence.tables import metadata

config = context.config
target_metadata = metadata


def include_object(obj, name, type_, reflected, compare_to):
    return name in target_metadata.tables if type_ == "table" else True


connection = config.attributes.get("connection")
if connection is None:
    raise RuntimeError("Run invoicing's migrations through slim_invoicing.persistence.migrate with a connection.")
context.configure(
    connection=connection,
    target_metadata=target_metadata,
    include_object=include_object,
    version_table=VERSION_TABLE,
    render_as_batch=True,  # SQLite needs batch mode for ALTERs
)
with context.begin_transaction():
    context.run_migrations()
