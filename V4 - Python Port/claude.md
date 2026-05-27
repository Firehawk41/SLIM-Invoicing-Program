# CLAUDE.md — LabPlus V4 Project Context

This file is loaded automatically by Claude Code. Read it before writing
any code. It defines the architecture, translation strategy, and conventions
for this project.

---

## Current Work / Resume Point

### Completed this session

**End-to-end pipeline validated** — `slim run-single` exits cleanly against the test submission:

```
slim --verbose run-single \
    --file "samples/051526 ACME ENV LABS Results [2026-05-16].xlsx" \
    --output out.csv
# → Wrote 1 sales order to out.csv.
```

Three line items produced, all internally consistent with seed data:
| Line item | Qty | Rate | Notes |
|---|---|---|---|
| 36 Elements + 7 Anions — Next Day | 2 | $85.00 | W-001+W-002 merged; regular rate (qty 2 < bulk_min 10) |
| pH — Next Day | 1 | $18.00 | W-003; Q2 no bulk discount |
| 36 Elements + 7 Anions — Five Days | 1 | $55.00 | W-004; Q1 Five Days regular rate |

Fixes required to reach clean run (all root-cause fixes, no symptom patches):
1. `create_test_submission.py` — sample cells had `"X"` markers (pre-port workaround); replaced with actual form_analyses names (`"36 Elements"`, `"7 Anions"`, `"pH"`). Cleared E21 (additional-elements sub-header) to prevent analysis_id=1 being injected into pH-only samples. Fixed W-004 from `"7 Anions only"` (no standalone quote) to `"36 Elements + 7 Anions"` (hits Q1 Five Days).
2. `seed_db.py` — quote dates were hardcoded (`expiry = date(2026, 4, 30)`, which had expired); changed to `today.replace(year=today.year + 2)` so seeds remain valid on any run date.
3. DB reseed — `form_analyses` table existed but was empty because the DB had been seeded before the `form_analyses` port was written. Deleted `slim.db` and re-ran `seed_db.py`.

**`form_chemicals` port** — fully wired end-to-end:

| File | Change |
|---|---|
| `slim/domain/chemical/chemical_repository.py` | Added `_FormChemicalRow` ORM model (`form_chemicals` table, PK=`form_name`, FK→`chemicals.ID`). Added `select_all_form_names() → list[tuple[str,int]]`. Fixed `insert_form_name()` (was `NotImplementedError`). |
| `slim/domain/chemical/chemical_cache.py` | Wired `_build()` to call `repo.select_all_form_names()` and populate `_by_form_name`. Added `add_form_name_index()`. |
| `slim/domain/chemical/chemical_service.py` | Fixed `add_form_name()` (was `NotImplementedError`); now cache-checks for idempotency, calls `repo.insert_form_name()`, then updates cache index. |
| `tests/domain/chemical/test_chemical.py` | Replaced 4 stub tests (expected `NotImplementedError`) with 8 real tests covering miss, hit, case-insensitivity, idempotency, DB roundtrip, and cache boot from pre-seeded rows. |
| `tests/domain/tr/test_form_input_resolver.py` | Added `_FormChemicalRow` import. Renamed stale test. Added `test_resolve_chemical_returns_chemical_on_hit` and `test_resolve_chemical_hit_is_case_insensitive`. |
| `seed_db.py` | Added 4 `form_chemicals` rows including the critical `"Water"` → `"DI Water Blank"` mapping (see comment block in that file). |

**`form_analyses` port** — fully wired end-to-end:

| File | Change |
|---|---|
| `slim/domain/analysis/analysis_repository.py` | Added `_FormAnalysisRow` ORM model (`form_analyses` table, composite PK=`(form_name, analysis_id)`, FK→`analyses.ID`). Added `select_all_form_analyses() → list[tuple[str,int]]`. Added `insert_form_analysis()`. |
| `slim/domain/analysis/analysis_cache.py` | Added `_by_form_name: dict[str, list[int]]`. Wired `_build()` to populate it. Fixed `get_ids_by_form_name()` (was `NotImplementedError`). Added `form_name_exists()`, `form_analysis_exists()`, `add_form_analysis_index()`. |
| `slim/domain/analysis/analysis_service.py` | Stored `self._repo`. Fixed `get_ids_by_form_name()`. Added `form_name_exists()` and `add_form_analysis()` write method. |
| `tests/domain/analysis/test_analysis.py` | Added `_FormAnalysisRow` import. Replaced stub test (expected `NotImplementedError`) with 8 real tests. |
| `seed_db.py` | Added `_FormAnalysisRow` import. Added 7 `form_analyses` rows: 5 exact-name aliases + a `"36 Elements + 7 Anions"` bundle row mapping to IDs 1 and 2. |

**332 tests passing** (was 296 before this session). Pipeline validated end-to-end.

### What is next

- **`test_parsers.py`** — previously excluded because openpyxl wasn't installed; it is now (v3.1.5). Re-enable it and verify the parser tests pass against the real submission file.

### Decisions made this session that aren't obvious from the code

- **`_FormChemicalRow` / `_FormAnalysisRow` are standalone ORM models, not `relationship()` on their parent.**
  Both mirror `_FormCustomerRow` on `Customer` exactly — separate `__tablename__` classes with FK columns.
  A SQLAlchemy `relationship()` was considered but rejected to stay consistent with the established pattern.

- **`form_analyses` uses a composite PK `(form_name, analysis_id)`, not `form_name` alone.**
  One form-name can map to multiple analysis IDs (bundle: `"36 Elements + 7 Anions"` → IDs 1 and 2).
  `_by_form_name` is therefore `dict[str, list[int]]`, and idempotency is checked per `(form_name, analysis_id)` pair.

- **`"Water"` maps to `"DI Water Blank"` (not a dedicated "Water" chemical).**
  `TRSampleFormParser._get_form_chemical_name()` unconditionally returns the literal string
  `"Water"` for every sample row in a WATER-type form — it never reads the Excel cell.
  So the `form_chemicals` table needs a `"Water"` entry pointing to whatever water-matrix
  blank chemical represents the generic water sample. `"DI Water Blank"` (id=2) was chosen.
  **This mapping is production-critical** — its absence makes every WATER-form submission fail.
  See the comment block in `seed_db.py` for full details.

---

## What This Project Is

A Python port of a VBA invoice automation system for a laboratory LIMS.
The system reads completed testing-request Excel forms, matches requested
analyses against quoted prices in a database, and writes a CSV of sales
orders for import into accounting software.

The VBA source (V3) is the reference implementation. This port preserves
the architecture and translates the patterns — it does not redesign.

---

## Architecture

Five independent domain stacks, each following the same pattern:

```
Entity → Repository → Cache → Service
```

Domains: **Analysis**, **Chemical**, **Customer**, **Element**, **Quote**

Two additional domains with no cache (transactional, not reference data):

- **TR** — submission parsing stack (parsers, resolver, service)
- **Sales Order** — entity, repository, service

The full pipeline sits above the domain layer:

```
xlsx files on disk
      │
SubmissionLoader (validates filenames, filters by date)
      │
TRSubmissionService.build_from_file (sheet scan → parser → TRSubmission)
      │
SalesOrderPricingEngine (group samples → match quotes → PricedGroup list)
      │
SalesOrderLineItemBuilder (resolve analysis names → SalesOrderLineItem list)
      │
SalesOrderBuilder (assemble SalesOrder → persist via SalesOrderService)
      │
SalesOrderCsvWriter (context manager → one CSV row per line item)
```

Shared infrastructure: SQLAlchemy session (database), Python logging (logger).
Both are injected at the composition root — never instantiated inside domain
classes.

---

## Translation Strategy

### Entity → Pydantic model

VBA entities used `Initialize`, `LoadFromDict`, `RequireKey`, `AssertInit`,
and `SetInternalID`. All of this dissolves.

```python
from pydantic import BaseModel, ConfigDict

class Analysis(BaseModel):
    model_config = ConfigDict(frozen=True)  # immutable after construction
    id: int
    name: str
    description: str
```

- `frozen=True` replaces the sealed-ID contract
- `model_validate(row_dict)` replaces `LoadFromDict`
- `SetInternalID` disappears — construct a new object with the DB-assigned ID
- No partial construction, so no single-init guard needed

### Repository → thin SQLAlchemy wrapper

- Parameterised queries only — no string concatenation
- `SqlStr` helper gone — SQLAlchemy handles escaping
- Returns domain entities, not raw rows
- `select_all` called only by Cache at startup
- `insert` returns the new integer ID

```python
class AnalysisRepository:
    def __init__(self, session: Session):
        self._session = session

    def select_all(self) -> list[Analysis]:
        ...

    def insert(self, analysis: Analysis) -> int:
        ...
```

### Cache → plain dict-based class

The VBA cache existed because every ADODB read had real latency. SQLAlchemy's
session identity map handles per-object caching automatically. The cache layer
is retained as a startup-loaded dict for name and form-name lookups that
would otherwise require repeated queries.

```python
class AnalysisCache:
    def __init__(self, repo: AnalysisRepository):
        self._by_id: dict[int, Analysis] = {}
        self._by_name: dict[str, Analysis] = {}
        self._by_form_name: dict[str, list[int]] = {}  # Analysis: one name → many IDs (bundles)
        self._build(repo)

# Chemical cache uses a different shape — one name maps to exactly one chemical:
#   self._by_form_name: dict[str, int] = {}            # Chemical: one name → one ID
#
# The difference matters: Analysis bundles (e.g. "36 Elements + 7 Anions") map a
# single form-name to multiple analysis IDs, so AnalysisCache uses list[int].
# ChemicalCache always resolves to a single chemical, so it uses a plain int.
# Idempotency in add_form_analysis() is therefore per (form_name, analysis_id) pair,
# whereas add_form_name() (chemicals) is per form_name alone.
```

- `NormName` becomes `name.lower().strip()` inline
- No `RemoveAll` — just reassign the dict on rebuild

### Service → same structure, no ceremony

- Owns repo and cache internally — callers never touch them
- Read pattern: all reads go through cache, zero DB round-trips
- Write pattern: validate → repo write → construct sealed entity → cache upsert
- Validation raises `ValueError` — no VBA error numbers

```python
class AnalysisService:
    def __init__(self, session: Session):
        repo = AnalysisRepository(session)
        self._repo = repo
        self._cache = AnalysisCache(repo)
```

---

## What Disappears

These VBA patterns have no Python equivalent and must not be recreated:

- `Initialize` / `AssertInit` — use `__init__`
- `On Error GoTo EH` / `HandleError` — use `try/except` only where needed
- `RequireKey` — Pydantic validation handles this
- `SetInternalID` — construct a new frozen object with the real ID
- `SqlStr` — SQLAlchemy parameterised queries
- `NullToDefault` — SQLAlchemy handles NULL at the ORM boundary
- `DebugMode` gate — use Python logging levels and pytest
- `Class_Initialize` / `Class_Terminate` — use `__init__` / context managers
- `Friend` access — use leading underscore convention (`_set_internal_id`)
- `modUtilities` helpers — covered by enum properties (`.label`, `.days`, `.from_form_string`)

---

## What Survives

These patterns carry over directly:

- Entity → Repository → Cache → Service layering
- Service as the sole public API — callers never touch repo or cache
- Incremental cache updates after writes — **except Quote**, which does a full
  rebuild because the three-level object graph (Quote → LineItem → Price) is
  too complex to update incrementally
- Raises on unmatched lookups (`GetIDByName` raises if not found)
- Composition root injects all dependencies — nothing self-constructs
- `id = 0` means unsaved; positive ID means persisted

---

## Database

**Development:** SQLite (`slim.db` in project root, gitignored)
**Production target:** PostgreSQL

Engine is configured in `infrastructure/database.py` and injected everywhere.
URL resolution uses the following precedence (highest → lowest):

1. `--db` CLI flag passed directly to `make_engine(url)` — overrides everything
2. `DB_URL` environment variable (loaded from `.env` at import time via python-dotenv)
3. Hardcoded fallback: `"sqlite:///slim.db"`

Supported URL formats (same examples as `.env.template`):

```
# SQLite (dev)
DB_URL=sqlite:///slim.db

# PostgreSQL (prod)
DB_URL=postgresql+psycopg2://user:pass@host/labplus
```

Sensitive config (DB path, credentials) lives in `.env` (gitignored). Never hardcode.
Copy `.env.template` to `.env` and fill in values to override the default.

**Install modes:**

```bash
# SQLite only (default — no extra dependencies)
pip install -e .

# PostgreSQL support
pip install -e ".[postgres]"   # pulls in psycopg2-binary
```

---

## Conventions

- One file per class: `analysis.py`, `analysis_repository.py`, etc.
- Domain folder per stack: `slim/domain/analysis/`
- `__init__.py` in each folder exports the public API (service only)
- Tests mirror the source structure: `tests/domain/analysis/`
- Every domain stack gets a test file before moving to the next domain
- Type hints everywhere
- Docstrings on public methods only
- **Always run tests as `python -m pytest`**, never bare `pytest`. The bare
  `pytest` binary at `/root/.local/bin/pytest` resolves to a different
  environment and fails with `ModuleNotFoundError: No module named 'openpyxl'`
  on import of `tests/domain/tr/test_parsers.py`.
- **`.env.template`** lives in the project root. Copy it to `.env` and fill in
  values before running the app. This follows the same "template file tracked,
  live file gitignored" pattern as V3's `modConfig.template.bas`.

### Settled type mappings

| VBA type | Python type | Notes |
|---|---|---|
| `Currency` | `Decimal` | `from decimal import Decimal`; `Numeric(10,4,asdecimal=True)` in ORM |
| `Collection` of IDs | `tuple[int, ...]` | Ordered, immutable |
| unordered set of IDs | `frozenset[int]` | e.g. KED element IDs on Chemical |
| nullable FK | `int \| None` | `nullable=True` on mapped column |
| optional `Date` | `date \| None` | e.g. `modified_date` on Quote |
| `PaymentTypeEnum` | `PaymentType(IntEnum)` | `PO_NUMBER=1`, `CREDIT_CARD=2` — in `slim/domain/quote/enums.py` |
| `ProcessingTimeEnum` | `ProcessingTime(IntEnum)` | in `slim/domain/tr/enums.py` |
| `TestingRequestTypeEnum` | `RequestType(IntEnum)` | same file; `CHEMICAL=1`, `WATER=2`, `WAFER=3` |
| pipeline working DTO (mutable) | `@dataclass` (not frozen) | e.g. `_GroupAccumulator` inside pricing engine |
| pipeline stage contract (immutable) | `@dataclass(frozen=True)` | e.g. `PricedGroup`, `LineItemData` |

### Sentinel values

- `SENTINEL_DATE = date(1901, 1, 1)` — missing service date in TR submissions
- `date.min` (`date(1,1,1)`) — "unset" date in `Quote.validate()`
- `Decimal("-1")` — no bulk pricing on `QuotePrice`

### TR domain notes

- `filename` parameter on `build_from_worksheet` — openpyxl rejects `[`/`]` in
  tab titles; service date is extracted from the filename instead
- `form_chemicals` and `form_analyses` are **both fully ported** (see `## Current Work / Resume Point` above)
- `TRFormInputResolver` is fast-path only: raises `UnresolvedCustomerError` /
  `UnresolvedChemicalError` on cache miss (no VBA UI slow-path equivalent)
- `build_from_file` scans all worksheets for the one whose D3 cell contains
  `"Testing Request Form"` — it does **not** use `wb.active`

### Quote domain notes

- Default-table ORM rows alias their FK columns to the same Python attribute
  names as the customer-table rows (e.g. `default_quote_id` → `.quote_id`) so
  the cache helpers work uniformly across both table sets
- `save_quote` updates the quote header only — line items require delete + recreate
- `ChemicalName` was dropped from `QuoteLineItem` (entity stays pure)

### Sales Order domain notes

- `SalesOrder` and `SalesOrderLineItem` are frozen Pydantic models
- No cache — sales orders are transactional records, not reference data
- `SalesOrderRepository.insert` returns a fully sealed entity: header gets the
  DB-assigned ID and each line item gets its own ID via `model_copy`
- `save` (update) touches the header only — line items are not modified

### Pipeline notes (`slim/pipeline/`)

- `PricedGroup` (frozen dataclass) — pricing engine output; has `frozenset[int]`
  analysis IDs, no display strings
- `LineItemData` (frozen dataclass) — builder's enriched intermediate; has
  item/description strings, no analysis IDs
- `SalesOrderPricingEngine.calculate_pricing_information` accepts
  `today: date | None = None` for testability (defaults to `date.today()`)
- Pricing: customer quotes tried first, default quotes as fallback; quotes sorted
  so bundles (more analyses per line item) are matched before singles
- Merge step: groups sharing the same `quote_price_id` are collapsed into one
  `PricedGroup` by summing quantities
- `ProcessingTime.label` provides display strings (e.g. `"Next Day"`) — no
  separate `ProcessingTimeToString` utility needed
- `SalesOrderCsvWriter` is a context manager:
  `with SalesOrderCsvWriter(path) as w: w.write(so)`
- `SubmissionLoader` validates filenames before opening any file:
  - Extension must be `.xlsx`
  - Filename must contain `"Results"` and must not contain `"Partial"`
  - First 6 characters must be a numeric MMDDYY date
- `load_by_date_range` isolates per-file errors: one bad file is logged and
  skipped; the rest of the batch continues

---

## Build Order

All domains and pipeline stages are complete.

1. **Element** — no dependencies ✓
2. **Analysis** — no dependencies ✓
3. **Chemical** — depends on Element ✓
4. **Customer** — no dependencies ✓
5. **TR** — submission parsing stack (parsers, enums, resolver, service) ✓
6. **Quote** — depends on Analysis, Chemical, Customer ✓
7. **SO Pipeline** — depends on all domains ✓
8. **Input Pipeline** — file validator, submission loader ✓
9. **Composition root + CLI** — `slim/app.py`, `slim/cli.py` ✓

---

## Current State

**332 tests passing.** Port is feature-complete against the VBA V3 reference.

| Stack | Tests | Notes |
|---|---|---|
| Element | 12 | no dependencies |
| Analysis | 17 | includes form_analyses tests |
| Chemical | 28 | depends on Element; includes form_chemicals tests |
| Customer | 33 | no dependencies |
| TR | 95 | parsers (incl. test_parsers.py), entities, enums, resolver |
| Quote | 46 | QuotePrice → QuoteLineItem → Quote hierarchy |
| Sales Order | 32 | domain entities, repo, service |
| Pipeline | 69 | pricing engine, builders, CSV writer, file validator, submission loader |

### Composition root and CLI

- **`slim/app.py`** — `create_app(session) -> App` factory; wires all domain services
  and pipeline stages. `App` is a frozen dataclass with `loader` and `builder` fields.
- **`slim/cli.py`** — Click CLI registered as the `slim` entry point:
  - `slim run --folder … --start … --end … --output …` — batch date-range mode
  - `slim run-single --file … --output …` — single file mode
  - Both call `Base.metadata.create_all(engine)` on startup and accept `--db` / `--verbose`

### What is not ported

Nothing outstanding. All domain stacks, form lookup tables, and parser tests are
fully ported and passing. `test_parsers.py` (23 tests) runs under `python -m pytest`
and was always included in the 332 total; it only fails under the bare `pytest`
binary due to an environment mismatch (see Conventions above).

### TODO

Nothing outstanding.
