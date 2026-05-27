import os

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session

# Load .env from the project root (or any parent directory) at import time.
# Variables already present in the environment take precedence over .env values.
load_dotenv()

# Precedence: DB_URL in .env (or real environment) > hardcoded SQLite default.
# The --db CLI flag sits above both of these — it passes a value to make_engine()
# directly, bypassing _DEFAULT_DB_URL entirely.
_DEFAULT_DB_URL: str = os.environ.get("DB_URL", "sqlite:///slim.db")


class Base(DeclarativeBase):
    pass


def make_engine(url: str | None = None):
    """Return a SQLAlchemy engine.

    Precedence: url argument (--db CLI flag) > DB_URL env var (.env) > sqlite:///slim.db
    """
    return create_engine(url or _DEFAULT_DB_URL)


def make_session(engine) -> Session:
    return Session(engine)
