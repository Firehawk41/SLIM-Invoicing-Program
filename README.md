# SLIM Invoice System

A laboratory invoicing automation system built across four iterations over several years — each version a direct response to a concrete limitation in the one before it.

The progression from a single procedural module to a fully layered Python application with 355 tests isn't academic. Every architectural decision in the final version was motivated by something that could be done better.

---

## The Short Version

V1 worked. V2 introduced classes but coupled them to the output format — when the output requirement changed in production, the refactor was substantial. V3 fixed that with proper separation of concerns. V4 ports the proven architecture to Python, replacing VBA ceremony with modern idioms without redesigning anything that was already right.

---

## What It Does

Reads completed testing-request Excel forms submitted by lab customers, matches the requested analyses against quoted prices in a database, and writes a CSV of sales orders ready for import into accounting software.

---

## Navigating This Repo

| Path | What's there |
|---|---|
| `V1/` | Procedural VBA — the starting point |
| `V2/` | Class-based VBA — introduced structure, exposed the coupling problem |
| `V3/` | Production VBA — clean domain/pipeline separation, tested against real data |
| `V4/` | Python port — SQLAlchemy, Pydantic, Click, 355 tests, end-to-end validated |

Start with `V4/README.md` for the full architecture and design narrative. The earlier versions are there to show the reasoning, not as reference implementations.

---

## Key Design Principles

- **Domain objects are shaped by their domain, not their output format** — the limitation that motivated V3
- **Separation of concerns** — domain layer, pipeline layer, and UI layer are independent
- **Single responsibility** — each class does one thing
- **Composition root** — all dependencies constructed once and injected; nothing self-constructs
- **Testability** — services and repositories have clean contracts with no hidden dependencies

---

## Technical Stack (V4)

Python 3.11 · SQLAlchemy · Pydantic · Click · pytest · SQLite (dev) · PostgreSQL (prod target)
