"""seed_db.py — populate slim.db with synthetic but internally consistent data.

Run from the project root:
    python seed_db.py

What gets created
-----------------
Elements       : Lead, Arsenic, Cadmium, Mercury
Analyses       : 36 Elements, 7 Anions, Titration, pH, ICP-MS Metals
Chemicals      : Acid Digestion Mix  (KED elements: Pb, As)
                 DI Water Blank      (no KED elements)
                 HNO3 Standard       (KED elements: Cd, Hg)
FormChemicals  : "Acid Digestion Mix" / "DI Water Blank" / "HNO3 Standard" (exact-name aliases)
                 "Water" → DI Water Blank  (WATER-form constant emitted by the parser)
Customers      : Acme Environmental Labs  → customer-specific quote
                 Pacific Water District   → uses default quote
Quotes         : Q1 (customer-specific, Acme, bundle: 36 Elements + 7 Anions, bulk pricing)
                 Q2 (customer-specific, Acme, single: pH)
                 DQ1 (default quote, covers Titration + ICP-MS Metals, bulk pricing)
Sales orders   : SO1 (Acme, references Q1 price rows)
"""

import sys
from datetime import date, datetime
from decimal import Decimal

# ---------------------------------------------------------------------------
# Bootstrap — make sure all ORM models are registered with Base.metadata
# before create_all is called.  The repository modules are private (_Foo),
# so we import each module; the class-level declarations fire on import.
# ---------------------------------------------------------------------------

import slim.domain.analysis.analysis_repository      # noqa: F401
import slim.domain.chemical.chemical_repository      # noqa: F401
import slim.domain.customer.customer_repository      # noqa: F401
import slim.domain.element.element_repository        # noqa: F401
import slim.domain.quote.quote_repository            # noqa: F401
import slim.domain.sales_order.sales_order_repository  # noqa: F401

from infrastructure.database import Base, make_engine, make_session

# Pull the private row classes out by name so we can build rows directly
# (avoids having to instantiate full domain-object graphs for seed data).
from slim.domain.analysis.analysis_repository import _AnalysisRow, _FormAnalysisRow
from slim.domain.chemical.chemical_repository import _ChemicalRow, _FormChemicalRow, _KEDElementRow
from slim.domain.customer.customer_repository import _CustomerRow, _FormCustomerRow
from slim.domain.element.element_repository import _ElementRow
from slim.domain.quote.quote_repository import (
    _DefaultQuoteAnalysisRow,
    _DefaultQuoteLineItemElementRow,
    _DefaultQuoteLineItemRow,
    _DefaultQuotePriceRow,
    _DefaultQuoteRow,
    _QuoteAnalysisRow,
    _QuoteCustomerRow,
    _QuoteLineItemElementRow,
    _QuoteLineItemRow,
    _QuotePriceRow,
    _QuoteRow,
)
from slim.domain.sales_order.sales_order_repository import (
    _SalesOrderLineItemRow,
    _SalesOrderRow,
)


def banner(msg: str) -> None:
    print(f"\n{'─' * 60}\n  {msg}\n{'─' * 60}")


def ok(label: str, obj_id: int | str) -> None:
    print(f"  ✓  {label:<42} id={obj_id}")


# ===========================================================================
# Main
# ===========================================================================

def main() -> None:
    engine = make_engine()          # sqlite:///slim.db (project root)
    Base.metadata.create_all(engine)
    print("✓  create_all: all tables ensured")

    session = make_session(engine)

    # -----------------------------------------------------------------------
    # 1. Elements
    # -----------------------------------------------------------------------
    banner("Elements")

    elements_data = [
        (1, "Pb", "Lead"),
        (2, "As", "Arsenic"),
        (3, "Cd", "Cadmium"),
        (4, "Hg", "Mercury"),
    ]
    el_rows: dict[str, _ElementRow] = {}
    for eid, sym, name in elements_data:
        row = _ElementRow(id=eid, element_symbol=sym, element_name=name)
        session.add(row)
        session.flush()
        el_rows[sym] = row
        ok(f"Element  {sym} ({name})", row.id)

    session.commit()

    # -----------------------------------------------------------------------
    # 2. Analyses
    # -----------------------------------------------------------------------
    banner("Analyses")

    analyses_data = [
        (1, "36 Elements",   "Full ICP-MS sweep of 36 target elements"),
        (2, "7 Anions",      "Ion chromatography for 7 common anions"),
        (3, "Titration",     "Acid/base titration for alkalinity"),
        (4, "pH",            "Direct pH measurement"),
        (5, "ICP-MS Metals", "Trace metals by ICP-MS, 200-series method"),
    ]
    an_rows: dict[str, _AnalysisRow] = {}
    for aid, name, desc in analyses_data:
        row = _AnalysisRow(id=aid, analysis_name=name, analysis_description=desc)
        session.add(row)
        session.flush()
        an_rows[name] = row
        ok(f"Analysis {name}", row.id)

    session.commit()

    # -----------------------------------------------------------------------
    # 3. Chemicals  (IDs are auto-assigned; flush to materialise)
    # -----------------------------------------------------------------------
    banner("Chemicals")

    chem_data = [
        # (name, metals_prep, silicon_prep, ions_prep, ked_syms)
        ("Acid Digestion Mix",  "HNO3/HCl digest",   "HF trace",    "dilute HNO3", ["Pb", "As"]),
        ("DI Water Blank",      "none",               "none",        "none",         []),
        ("HNO3 Standard",       "2% HNO3 dilution",  "none",        "filtered",    ["Cd", "Hg"]),
    ]
    chem_rows: dict[str, _ChemicalRow] = {}
    for cname, metals, silicon, ions, ked_syms in chem_data:
        chem = _ChemicalRow(
            chemical_name=cname,
            metals_prep=metals,
            silicon_prep=silicon,
            ions_prep=ions,
            database_entry_date=datetime(2025, 1, 15, 9, 0, 0),
        )
        session.add(chem)
        session.flush()  # get chem.id before inserting KED children
        for sym in ked_syms:
            session.add(_KEDElementRow(chemical_id=chem.id, element_id=el_rows[sym].id))
        session.flush()
        chem_rows[cname] = chem
        ok(f"Chemical {cname!r}  KED={ked_syms or '—'}", chem.id)

    # form_chemicals: map Excel cell strings → chemical IDs.
    #
    # Every form_chemicals row answers the question:
    #   "What does this text string in an Excel cell map to in the chemicals table?"
    #
    # CHEMICAL forms:  TRSampleFormParser reads the raw cell value from column 3
    #                  of each sample row and passes it to resolve_chemical().
    # WATER forms:     TRSampleFormParser._get_form_chemical_name() IGNORES the
    #                  Excel cell entirely and unconditionally returns the literal
    #                  string "Water" for every sample row.
    #
    # ⚠️  PRODUCTION CRITICAL — "Water" → "DI Water Blank" mapping:
    #     The parser for WATER-type testing-request forms hard-codes the string
    #     "Water" as the chemical name for every sample, regardless of what
    #     chemical was actually used in the lab.  If this row is absent from
    #     form_chemicals, TRFormInputResolver.resolve_chemical("Water") raises
    #     UnresolvedChemicalError and the ENTIRE water-form pipeline fails —
    #     no samples from any WATER submission will be invoiced.
    #
    #     "Water" must always point to the water-matrix blank chemical (here:
    #     "DI Water Blank", id=2).  Do not delete or rename this row.
    form_chemical_data = [
        ("Acid Digestion Mix",  "Acid Digestion Mix"),  # CHEMICAL form col-3 value
        ("DI Water Blank",      "DI Water Blank"),       # CHEMICAL / WAFER exact-name alias
        ("HNO3 Standard",       "HNO3 Standard"),        # CHEMICAL / WAFER exact-name alias
        # ↓ CRITICAL: WATER forms emit this literal string; must map to a water-matrix chemical
        ("Water",               "DI Water Blank"),
    ]
    for form_name, chem_name in form_chemical_data:
        session.add(_FormChemicalRow(
            form_name=form_name,
            chemical_id=chem_rows[chem_name].id,
        ))
    session.commit()
    for form_name, chem_name in form_chemical_data:
        ok(f"FormChemical {form_name!r} → {chem_name!r}", chem_rows[chem_name].id)

    # form_analyses: map Excel cell strings → analysis IDs.
    #
    # Unlike form_chemicals (1-to-1), form_analyses is 1-to-many: one form-name
    # key can map to multiple analysis IDs (e.g. a bundle test name that covers
    # several analyses in one line item).
    #
    # The form_name here is the text that appears in the Excel column header of
    # the testing-request form (one column per analysis block).  It is passed to
    # AnalysisService.get_ids_by_form_name() and must match case-insensitively.
    form_analysis_data = [
        # (form_name,                  analysis_name)   — 1-to-1 exact aliases
        ("36 Elements",               "36 Elements"),
        ("7 Anions",                  "7 Anions"),
        ("Titration",                 "Titration"),
        ("pH",                        "pH"),
        ("ICP-MS Metals",             "ICP-MS Metals"),
        # bundle: one Excel column header → two analysis IDs
        ("36 Elements + 7 Anions",    "36 Elements"),
        ("36 Elements + 7 Anions",    "7 Anions"),
    ]
    for form_name, analysis_name in form_analysis_data:
        session.add(_FormAnalysisRow(
            form_name=form_name,
            analysis_id=an_rows[analysis_name].id,
        ))
    session.commit()
    # Print unique form names with their resolved analysis IDs
    seen: dict[str, list[int]] = {}
    for form_name, analysis_name in form_analysis_data:
        seen.setdefault(form_name, []).append(an_rows[analysis_name].id)
    for form_name, ids in seen.items():
        ok(f"FormAnalysis {form_name!r} → ids={ids}", ids[0])

    # -----------------------------------------------------------------------
    # 4. Customers
    # -----------------------------------------------------------------------
    banner("Customers")

    acme = _CustomerRow(
        customer_name="Acme Environmental Labs",
        street_address="1200 Industrial Blvd",
        city="Sacramento",
        state="CA",
        postal_code="95814",
        country="USA",
    )
    pacific = _CustomerRow(
        customer_name="Pacific Water District",
        street_address="500 Reservoir Rd",
        city="Oakland",
        state="CA",
        postal_code="94612",
        country="USA",
    )
    session.add_all([acme, pacific])
    session.flush()
    ok("Customer Acme Environmental Labs", acme.id)
    ok("Customer Pacific Water District",  pacific.id)

    # Form-customer aliases
    session.add(_FormCustomerRow(
        form_customer="ACME ENV LABS",
        form_address="1200 Industrial Blvd",
        form_address_2="Sacramento CA 95814",
        customer_id=acme.id,
    ))
    session.add(_FormCustomerRow(
        form_customer="PACIFIC WATER DIST",
        form_address="500 Reservoir Rd",
        form_address_2="Oakland CA 94612",
        customer_id=pacific.id,
    ))
    session.commit()
    print("  ✓  form_customer aliases added")

    # -----------------------------------------------------------------------
    # 5. Customer-specific quotes  (Acme)
    #
    #   Q1 — bundle: 36 Elements + 7 Anions
    #          Two processing tiers: Next Day (id=2) and Five Days (id=8)
    #          bulk_min=10 so bulk pricing kicks in at 10+ samples
    #
    #   Q2 — single analysis: pH
    #          One tier: Next Day (id=2), no bulk discount
    # -----------------------------------------------------------------------
    banner("Customer-Specific Quotes  (Acme)")

    # Use relative dates so the seed quotes remain valid regardless of when the
    # script is run.  Effective from 60 days ago, expires in 2 years.
    today = date.today()
    expiry = today.replace(year=today.year + 2)

    # --- Q1: bundle quote with bulk pricing ---
    q1 = _QuoteRow(
        effective_date=today,
        expiry_date=expiry,
        payment_type=1,   # PO_NUMBER
        payment_value=30, # net-30
        valid=True,
        created_by="seed_db",
        created_date=today,
        modified_date=None,
        comments="Bundle quote for 36 Elements + 7 Anions — Acme",
    )
    session.add(q1)
    session.flush()
    ok("Quote Q1 (bundle, bulk pricing)", q1.id)

    # Q1 line item: CHEMICAL request, uses Acid Digestion Mix, bulk_min=10
    q1_li = _QuoteLineItemRow(
        quote_id=q1.id,
        request_type_id=1,   # CHEMICAL
        chemical_id=chem_rows["Acid Digestion Mix"].id,
        bulk_min=10,
        valid=True,
    )
    session.add(q1_li)
    session.flush()
    ok("  └─ Q1 line item", q1_li.id)

    # Q1 prices: Next Day tier (id=2) — regular $85, bulk $68
    #            Five Days tier (id=8) — regular $55, bulk $44
    q1_price_nd = _QuotePriceRow(
        quote_line_item_id=q1_li.id,
        processing_time_id=2,             # NEXT_DAY
        regular_price=Decimal("85.0000"),
        bulk_price=Decimal("68.0000"),
        valid=True,
    )
    q1_price_5d = _QuotePriceRow(
        quote_line_item_id=q1_li.id,
        processing_time_id=8,             # FIVE_DAYS
        regular_price=Decimal("55.0000"),
        bulk_price=Decimal("44.0000"),
        valid=True,
    )
    session.add_all([q1_price_nd, q1_price_5d])
    session.flush()
    ok("  └─ Q1 price (Next Day  reg=$85 bulk=$68)", q1_price_nd.id)
    ok("  └─ Q1 price (Five Days reg=$55 bulk=$44)", q1_price_5d.id)

    # Q1 analyses: 36 Elements (id=1) + 7 Anions (id=2)
    session.add(_QuoteAnalysisRow(quote_line_item_id=q1_li.id, analysis_id=an_rows["36 Elements"].id))
    session.add(_QuoteAnalysisRow(quote_line_item_id=q1_li.id, analysis_id=an_rows["7 Anions"].id))

    # Q1 elements: Pb, As (the KED elements from Acid Digestion Mix)
    session.add(_QuoteLineItemElementRow(quote_line_item_id=q1_li.id, element_id=el_rows["Pb"].id))
    session.add(_QuoteLineItemElementRow(quote_line_item_id=q1_li.id, element_id=el_rows["As"].id))

    # Link Q1 → Acme
    session.add(_QuoteCustomerRow(quote_id=q1.id, customer_id=acme.id))

    session.commit()
    print("  ✓  Q1 analyses + elements + customer link committed")

    # --- Q2: single-analysis quote, pH ---
    q2 = _QuoteRow(
        effective_date=today,
        expiry_date=expiry,
        payment_type=1,   # PO_NUMBER
        payment_value=30,
        valid=True,
        created_by="seed_db",
        created_date=today,
        modified_date=None,
        comments="Single-analysis pH quote — Acme",
    )
    session.add(q2)
    session.flush()
    ok("Quote Q2 (single: pH)", q2.id)

    q2_li = _QuoteLineItemRow(
        quote_id=q2.id,
        request_type_id=2,   # WATER
        chemical_id=None,    # no specific chemical for pH
        bulk_min=0,
        valid=True,
    )
    session.add(q2_li)
    session.flush()
    ok("  └─ Q2 line item", q2_li.id)

    q2_price = _QuotePriceRow(
        quote_line_item_id=q2_li.id,
        processing_time_id=2,             # NEXT_DAY
        regular_price=Decimal("18.0000"),
        bulk_price=Decimal("18.0000"),    # no discount — same as regular
        valid=True,
    )
    session.add(q2_price)
    session.flush()
    ok("  └─ Q2 price (Next Day  reg=$18 bulk=$18)", q2_price.id)

    session.add(_QuoteAnalysisRow(quote_line_item_id=q2_li.id, analysis_id=an_rows["pH"].id))
    session.add(_QuoteCustomerRow(quote_id=q2.id, customer_id=acme.id))
    session.commit()
    print("  ✓  Q2 analysis + customer link committed")

    # -----------------------------------------------------------------------
    # 6. Default quote  (Pacific Water District uses this)
    #
    #   DQ1 — bundle: Titration + ICP-MS Metals
    #          Two tiers: Next Day and Three Days
    #          bulk_min=5 — meaningful bulk discount
    # -----------------------------------------------------------------------
    banner("Default Quote  (Pacific Water District)")

    dq1 = _DefaultQuoteRow(
        effective_date=today,
        expiry_date=expiry,
        payment_type=2,   # CREDIT_CARD
        payment_value=0,  # credit card = no net terms
        valid=True,
        created_by="seed_db",
        created_date=today,
        modified_date=None,
        comments="Default rate sheet — Titration + ICP-MS Metals",
    )
    session.add(dq1)
    session.flush()
    ok("Default Quote DQ1 (bundle, bulk pricing)", dq1.id)

    dq1_li = _DefaultQuoteLineItemRow(
        quote_id=dq1.id,
        request_type_id=2,   # WATER
        chemical_id=None,
        bulk_min=5,
        valid=True,
    )
    session.add(dq1_li)
    session.flush()
    ok("  └─ DQ1 line item", dq1_li.id)

    dq1_price_nd = _DefaultQuotePriceRow(
        quote_line_item_id=dq1_li.id,
        processing_time_id=2,              # NEXT_DAY
        regular_price=Decimal("120.0000"),
        bulk_price=Decimal("96.0000"),
        valid=True,
    )
    dq1_price_3d = _DefaultQuotePriceRow(
        quote_line_item_id=dq1_li.id,
        processing_time_id=7,              # THREE_DAYS
        regular_price=Decimal("75.0000"),
        bulk_price=Decimal("60.0000"),
        valid=True,
    )
    session.add_all([dq1_price_nd, dq1_price_3d])
    session.flush()
    ok("  └─ DQ1 price (Next Day   reg=$120 bulk=$96)", dq1_price_nd.id)
    ok("  └─ DQ1 price (Three Days reg=$75  bulk=$60)", dq1_price_3d.id)

    session.add(_DefaultQuoteAnalysisRow(
        quote_line_item_id=dq1_li.id, analysis_id=an_rows["Titration"].id))
    session.add(_DefaultQuoteAnalysisRow(
        quote_line_item_id=dq1_li.id, analysis_id=an_rows["ICP-MS Metals"].id))

    # Elements associated with ICP-MS Metals: Cd, Hg (from HNO3 Standard)
    session.add(_DefaultQuoteLineItemElementRow(
        quote_line_item_id=dq1_li.id, element_id=el_rows["Cd"].id))
    session.add(_DefaultQuoteLineItemElementRow(
        quote_line_item_id=dq1_li.id, element_id=el_rows["Hg"].id))

    session.commit()
    print("  ✓  DQ1 analyses + elements committed")

    # -----------------------------------------------------------------------
    # 7. Sales order  (Acme, Q1 Next-Day pricing, 12 samples → bulk rate)
    # -----------------------------------------------------------------------
    banner("Sales Order")

    so = _SalesOrderRow(
        customer_name="Acme Environmental Labs",
        customer_id=acme.id,
        submission_id=10001,
        submission_ref="ACE-2025-001",
        date_received=date(2025, 5, 5),
        po_number="PO-88421",
        payment_type=1,   # PO_NUMBER
        payment_terms=30,
        requester_name="Dr. Jane Carter",
        requester_email="jcarter@acmelabs.example.com",
        service_date=date(2025, 5, 6),
    )
    session.add(so)
    session.flush()
    ok("Sales Order SO1", so.id)

    # 12 samples at bulk rate $68 (Next Day, Q1)
    so_li = _SalesOrderLineItemRow(
        sales_order_id=so.id,
        item="36 Elements + 7 Anions (Next Day)",
        description="ICP-MS 36-element suite + IC 7-anion panel, next-day turnaround",
        quantity=12,
        rate=Decimal("68.0000"),
        quote_price_id=q1_price_nd.id,
        processing_time_id=2,             # NEXT_DAY
    )
    session.add(so_li)
    session.flush()
    ok("  └─ SO1 line item (qty=12 × $68.00 = $816.00)", so_li.id)

    session.commit()

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    banner("Seed complete — record counts")
    for table in [
        "elements", "analyses", "chemicals", "ked_elements", "form_chemicals",
        "customers", "form_customers",
        "quotes", "quote_line_items", "quote_prices",
        "quote_analysis", "quote_line_item_element", "quote_customer",
        "default_quotes", "default_quote_line_items", "default_quote_prices",
        "default_quote_analysis", "default_quote_line_item_element",
        "sales_orders", "sales_order_items",
    ]:
        from sqlalchemy import text as _text
        count = session.execute(_text(f"SELECT COUNT(*) FROM {table}")).scalar()
        print(f"  {table:<42} {count:>3} row(s)")

    session.close()
    print("\n✓  Done.  Database: slim.db\n")


if __name__ == "__main__":
    main()
