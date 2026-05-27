# SLIM Invoice System — V3

A VBA/Excel automation system that reads completed lab testing-request forms,
matches requested analyses against quoted prices in an Access database, and
writes a CSV of sales orders ready for import into accounting software.

Built and tested against real production data at a laboratory LIMS.

> **This is the V3 reference implementation.** V4 (Python port) is complete
> and is the current version. See `V4/README.md` for the full architecture
> and design narrative.

---

## Why This Exists — The Design Progression

This is the third iteration of the same system. Each version is a direct
response to a concrete limitation in the one before it.

**V1 — Procedural**
A single module: loops, arrays, no structure. Functional, but every new
requirement meant editing everything. Adding a customer field required
touching the whole file.

**V2 — Classes, But Coupled**
Introduced classes, but invoice objects were shaped around their output
format (Word). When the output requirement changed to CSV in production,
the refactor was substantial — the domain objects had no independence from
the writer. This is the limitation that motivated V3.

**V3 — This Codebase**
Clean separation between domain objects and output. The CSV writer is a
thin adapter that reads from domain objects — it has no influence on their
structure. Swapping to a new output format (Excel, PDF, API) requires only
a new writer class. The domain layer doesn't change.

**V4 — Python Port (complete)**
The architecture ports directly. Domain layer, service layer, composition
root, and pipeline separation all survive the language change. The VBA
ceremony (single-init guards, dict hydration, manual SQL) dissolves into
Python idioms — SQLAlchemy, Pydantic, Click. 355 tests, end-to-end
validated. See `V4/README.md`.

---

## Architecture

The codebase has two distinct layers.

### Domain Layer

Four independent domains each follow the same four-layer pattern:

```
Entity → Repository → Cache → Service
```

Domains: **Analysis**, **Chemical**, **Customer**, **Element**, **Quote**

All reads go through an in-memory cache loaded at startup. Writes hit the
database and update the cache incrementally. Domain services are pure —
no UI coupling.

### Sales Order Pipeline

Sales orders are transactional records — no cache. The pipeline runs above
the domain layer:

```
Submission Parsing → Pricing → Line Item Building → SO Building → Persistence → CSV Output
```

Key pipeline stages:

* **`clsInvoiceSubmissionManager`** — scans a folder of `.xlsx` submission
  forms, validates filenames, filters by date range, builds submission objects
* **`clsSalesOrderPricingEngine`** — groups samples by analysis set and
  processing time, matches against customer-specific or default quotes,
  merges identical line items by `QuotePriceID`
* **`clsSalesOrderBuilder`** — coordinates the line item builder and service,
  constructs `clsSalesOrder` entities ready for persistence
* **`clsSalesOrderWriterCSV`** — thin output adapter; reads only from public
  properties on domain objects, writes the CSV file

### Shared Infrastructure

* **`clsAccessDatabase`** — ADODB wrapper, injected everywhere
* **`clsLoggingSystem`** — structured logger, injected everywhere
* **`modInvoiceSystem`** — composition root; all services constructed once
  and injected, no business logic

---

## Folder Structure

```
V3/
├── Domain/         Entity, Repository, Cache, Service for each domain
├── Pipeline/       Sales order entities, builders, pricing engine, writer
├── Parsing/        Submission parsing and TR tree classes
├── Infrastructure/ Database wrapper and logging system
├── UI/             Invoice entry UserForm
└── System/         Entry point, enums, utilities, config template
```

---

## Key Design Decisions

**Two hydration patterns, deliberately.** Domain layer entities use
`LoadFromDict` — they face a genuine many-source problem and need decoupling
from their hydration source. SO pipeline entities use DTOs as a universal
contract for both creation and rehydration paths, keeping a single entry
point into each entity regardless of data source.

**Pricing engine raises on unmatched groups** rather than silently producing
zero-price line items. Unpriced output would be a silent data integrity
failure.

**`QuotePriceID` as merge key.** After pricing, identical line items (same
quote price) are collapsed into single items with accumulated quantities.
The same ID also provides traceability from any line item back to its source
quote.

**UI coupling contained to one class.** `clsTRFormInputResolver` is the
sole location where form field strings are resolved to domain objects.
Domain services have no awareness of the UI.

**`modConfig` excluded from source control.** Sensitive paths live in
`modConfig.bas`, which is `.gitignored`. `modConfig.template.bas` documents
the required constants with placeholder values.

---

## Intentionally Deferred

These classes are scoped out of V3 — not gaps, but deliberate boundaries:

| Class | Purpose |
|---|---|
| `clsSubmissionRepository` | DB persistence for parsed submissions |
| `clsSampleRepository` | DB persistence for individual samples |
| `clsCustomerBillingProfile` | Per-customer payment terms |
| `clsQuoteFormController` | UI controller for quote editor (separate project) |

Transaction coordination across the SO header and its line items is also
deferred — `clsAccessDatabase` does not expose explicit
`BeginTransaction`/`Commit`/`Rollback` control. Addressed in V4 via
SQLAlchemy session management.

---

## Setup

1. Copy `modConfig.template.bas` to `modConfig.bas`
2. Populate `DB_PATH`, `OUTPUT_PATH`, and `DEBUG_MODE` with local values
3. Import all `.bas`, `.cls`, and `.frm` files into an Excel workbook
4. Run `CreateInvoice` from `modInvoiceSystem`
