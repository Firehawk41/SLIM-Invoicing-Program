"""Invoicing keeps its own version table and only its own tables (D-83)."""

from sqlalchemy import inspect

from slim_invoicing.persistence import migrate
from slim_invoicing.persistence.tables import metadata


def test_upgrade_creates_exactly_the_declared_tables_and_downgrade_removes_them(engine):
    names = set(inspect(engine).get_table_names())
    assert names == set(metadata.tables) | {migrate.VERSION_TABLE}
    with engine.begin() as connection:
        migrate.downgrade(connection)
    assert set(inspect(engine).get_table_names()) == {migrate.VERSION_TABLE}
    with engine.begin() as connection:
        migrate.upgrade(connection)  # so the fixture's teardown has something to downgrade


def test_the_migration_matches_the_tables(engine):
    """Every column the code writes exists, as the migration created it."""
    inspector = inspect(engine)
    for table in metadata.sorted_tables:
        assert {c["name"] for c in inspector.get_columns(table.name)} == set(table.columns.keys())
