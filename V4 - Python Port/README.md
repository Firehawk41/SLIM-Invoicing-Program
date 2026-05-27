# SLIM Invoice System — V4 (Python Port)

A Python port of the VBA invoice automation system. Reads completed lab testing-request forms, matches requested analyses against quoted prices in a database, and writes a CSV of sales orders ready for import into accounting software.

This is the fourth iteration of the same system. The architecture is unchanged from V3 — the patterns port directly. The VBA ceremony dissolves.

---

## What Changed in Translation

V3 proved the architecture was right. V4 removes the language tax.

| V3 (VBA) | V4 (Python) |
|---|---|
| `Initialize` / `AssertInit` / single-init guard | `__init__` |
| `LoadFromDict` / `RequireKey` | `model_validate()` — Pydantic handles missing keys |
| `SetInternalID` / sealed-ID contract | `frozen=True` on Pydantic models — immutable after construction |
| `On Error GoTo EH` / `HandleError` | `try/except` only where needed |
| `SqlStr` / manual parameterisation | SQLAlchemy parameterised queries |
| `NullToDefault` at every DB boundary | SQLAlchemy handles NULL at the ORM boundary |
| `DebugMode` gate | Python logging levels and pytest |
| ADODB with real per-query latency | SQLAlchemy session identity map |
| Mutable DTO passed through pipeline stages | Two frozen dataclasses — `PricedGroup` and `LineItemData` — one per stage boundary |
| `Scripting.Dictionary` as a set | `frozenset` |
| `Currency` type | `Decimal` |
| Cross-object transaction coordination deferred | SQLAlchemy session handles this natively |

The domain layer (Entity → Repository → Cache → Service), the pipeline separation, the composition root, and the pricing logic are all direct translations. Nothing was redesigned.

---

## Architecture

### Domain Layer

Five independent domains, each following the same four-layer pattern:

```
Entity → Repository → Cache → Service
```

Domains: **Analysis**, **Chemical**, **Customer**, **Element**, **Quote**

Entities are frozen Pydantic models — immutable after construction. All reads go through an in-memory cache loaded at startup. Writes hit the database and update the cache incrementally. Services are the sole public API — callers never touch the repository or cache directly.

### Sales Order Pipeline

Sales orders are transactional records — no cache. The pipeline runs above the domain layer:

```
xlsx files on disk
      │
SubmissionLoader — validates filenames, filters by date range
      │
TRSubmissionService — sheet scan → parser → TRSubmission
      │
SalesOrderPricingEngine — group samples → match quotes → PricedGroup list
      │
SalesOrderLineItemBuilder — resolve analysis names → SalesOrderLineItem list
      │
SalesOrderBuilder — assemble SalesOrder → persist via SalesOrderService
      │
SalesOrderCsvWriter — context manager → one CSV row per line item
```

Key pipeline stages:

- **SubmissionLoader** — validates filenames before opening any file; isolates per-file errors so one bad file doesn't abort the batch
- **SalesOrderPricingEngine** — groups samples by analysis set and processing time, matches against customer-specific or default quotes, merges identical line items by `quote_price_id`; raises on unmatched groups rather than silently producing zero-price line items
- **SalesOrderBuilder** — coordinates the line item builder and service; constructs `SalesOrder` entities ready for persistence
- **SalesOrderCsvWriter** — thin output adapter; reads only from public properties on domain objects

### Shared Infrastructure

- `infrastructure/database.py` — SQLAlchemy engine; reads `DB_URL` from `.env`, falls back to `sqlite:///labplus.db`
- `labplus/app.py` — composition root; all services constructed once and injected, no business logic
- `labplus/cli.py` — Click CLI; `--db` flag overrides `.env` if provided

---

## Database

**Development:** SQLite (`labplus.db`, gitignored)  
**Production target:** PostgreSQL

To switch, set `DB_URL` in `.env`:

```
# SQLite (dev)
DB_URL=sqlite:///labplus.db

# PostgreSQL (prod)
DB_URL=postgresql+psycopg2://user:pass@host/labplus
```

PostgreSQL driver included as an optional dependency:

```bash
pip install -e ".[postgres]"
```

---

## Setup

```bash
git clone https://github.com/SLIM-Invoicing-Program/v4-Python-Port
cd v4-Python-Port
pip install -e .
```

Copy the environment template and configure:

```bash
cp .env.template .env
# edit .env if not using the SQLite default
```

Seed the database with synthetic reference data:

```bash
python seed_db.py
```

Run against the included test submission:

```bash
labplus run-single \
  --file "samples/051526 ACME ENV LABS Results [2026-05-16].xlsx" \
  --output out.csv --verbose
```

Or run in batch mode against a folder:

```bash
labplus run \
  --folder samples/ \
  --start 2026-05-01 \
  --end 2026-05-31 \
  --output out.csv
```

The test submission uses entirely synthetic data — fabricated customer name, sample IDs, and analysis requests. No real laboratory data is included in this repository.

---

## Tests

```bash
python -m pytest
```

355 tests. Use `python -m pytest`, not bare `pytest` — the project uses a virtual environment and the bare binary may resolve to a different Python.

---

## Folder Structure

```
labplus/
  domain/
    analysis/        Entity, Repository, Cache, Service
    chemical/        Entity, Repository, Cache, Service
    customer/        Entity, Repository, Cache, Service
    element/         Entity, Repository, Cache, Service
    quote/           Entity, Repository, Cache, Service
    sales_order/     Entity, Repository, Service (no cache — transactional)
    tr/              TRSubmission, TRSample, parser, resolver, service
  pipeline/          Pricing engine, builders, CSV writer
  infrastructure/    SQLAlchemy engine, Base, session factory
  app.py             Composition root
  cli.py             Click CLI entry point
samples/             Synthetic test submission file
tests/               Mirrors source structure
seed_db.py           Populates dev database with synthetic reference data
.env.template        Environment config template — copy to .env
```

---

## Key Design Decisions

**Pricing engine raises on unmatched groups.** An unmatched group would produce a zero-price line item and silently corrupt the invoice. The engine raises instead, surfacing the configuration gap immediately.

**`quote_price_id` as merge key.** After pricing, groups sharing the same quote price are collapsed into single line items with accumulated quantities. The same ID provides traceability from any line item back to its source quote.

**Two frozen dataclasses at pipeline stage boundaries.** `PricedGroup` is the pricing engine's output contract. `LineItemData` is the line item builder's output contract. Each stage produces a new immutable object — pipeline transformations are explicit and each stage is independently testable.

**UI coupling contained to one class.** `TRFormInputResolver` is the sole location where form strings are resolved to domain objects. Domain services have no awareness of any UI.

**`.env.template` committed, `.env` gitignored.** Follows the same pattern as V3's `modConfig.template.bas`. Sensitive config never touches source control.
