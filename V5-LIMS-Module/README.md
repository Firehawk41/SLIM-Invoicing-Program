# SLIM Invoice System — V5 (LIMS module)

V4 proved the architecture in Python. V5 moves it inside the LIMS: it bills what the lab **reported**,
straight from the LIMS, instead of reading Excel forms off a shared drive.

## What V4 couldn't do, and V5 does

| V4 | V5 |
|---|---|
| Read completed Testing Request `.xlsx` files from disk | The host app hands it a `BillableRequest`: the TR plus what its issued reports showed |
| Billed every analysis on the form | Bills each ordered analysis a report showed, once; never-reported work is never billed |
| Kept its own copies of customers, chemicals, analyses | Uses slim-domain's IDs: one set of reference data for the whole LIMS |
| A CLI run over a folder of files | A library: the host calls `invoice()` in its own transaction and moves the TR to invoiced in the same one |
| Nothing stopped a TR being billed twice | One sales order per TR, enforced by the service and by a unique key |

Ported unchanged in behaviour: quotes (customer quotes first, then defaults; widest line first, so a
bundle beats its parts; bulk price from the line's minimum), the pricing engine, and the accounting CSV
with V4's columns.

## Use (from the host)

```python
from slim_invoicing import BillableRequest, invoice, write_csv
from slim_invoicing.persistence import migrate

migrate.upgrade(connection)                  # its own tables (inv_*) and version table
order = invoice(session, request, "SUE")     # raises InvoiceRefused: incomplete, already invoiced,
                                             # nothing reported, or an analysis no quote prices
write_csv([order], file)                     # the accounting import file
```

Only a TR with a complete report is invoiced. Invoicing flushes; the host commits.

## Layout

```
slim_invoicing/
├── domain/       # enums (the host's codes), quote, sales_order, billable (the input, and the "bill once" rule)
├── pricing.py    # the V4 pricing engine
├── service.py    # invoice(): refuse, price, build, save
├── export.py     # the accounting CSV
└── persistence/  # tables (inv_*), repositories, Alembic (alembic_version_slim_invoicing)
```

## Develop

```bash
pip install -e ".[dev]"
pytest                                        # SQLite; TEST_DATABASE_URL=postgresql+psycopg2://... for Postgres
ruff check . && ruff format --check .
```

Decisions: slim-data-oss `docs/DECISIONS.md` D-43, D-83.
