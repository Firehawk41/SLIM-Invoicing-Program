import os

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy.pool import StaticPool

from slim_invoicing.persistence import migrate


@pytest.fixture
def engine():
    """SQLite in memory, or Postgres via TEST_DATABASE_URL (a demo database only)."""
    url = os.environ.get("TEST_DATABASE_URL")
    engine = create_engine(url) if url else create_engine("sqlite://", poolclass=StaticPool)
    with engine.begin() as connection:
        migrate.upgrade(connection)
    yield engine
    with engine.begin() as connection:
        migrate.downgrade(connection)
    engine.dispose()


@pytest.fixture
def session(engine):
    with Session(engine) as session:
        yield session
        session.rollback()
