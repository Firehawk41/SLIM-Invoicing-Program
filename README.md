# SLIM-Invoicing-Program
Invoicing Program for a laboratory across various iterations

## Quick Start

To understand the architecture:
1. Read this file (you're here)
2. Look at `V3/README.md` for the full design narrative
3. Browse `VBA_Codebase_Reference_V3.md` for class-by-class details

To run the system (requires Excel, VBA, Access database):
1. Open the workbook
2. Run `CreateInvoice()` macro
3. Select invoice mode (Individual file or Batch date range)
4. CSV is written to the configured output path

## Key Design Principles

- **Domain objects are shaped by their domain, not their output format**
- **Separation of concerns**: domain layer, pipeline layer, UI layer are independent
- **Single Responsibility**: each class does one thing well
- **Defensive programming**: all public methods validate preconditions
- **Testability**: services and repositories have clean contracts with no hidden dependencies

## V4 — Why Port to Python?

The VBA patterns work but have ceremony:
- Manual SQL instead of SQLAlchemy ORM
- Dict-based hydration instead of dataclasses/Pydantic
- Single-init guards instead of constructor idioms
- Manual transaction management instead of session contexts

The V4 port will keep the architecture (domain layer, service layer, composition root, pipeline separation) and dissolve the ceremony into Python idioms. The goal is to show that good architecture transcends language.

## Contact / Questions

This is a portfolio project. The code is complete and tested end-to-end. For technical deep dives, see `V3/README.md`.
