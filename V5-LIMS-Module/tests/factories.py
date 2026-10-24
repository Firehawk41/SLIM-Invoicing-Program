from datetime import date
from decimal import Decimal
from itertools import count

from slim_invoicing import (
    BillableRequest,
    BillableSample,
    PaymentMethod,
    ProcessingTime,
    Quote,
    QuoteLineItem,
    QuotePrice,
)

ELEMENTS, ANIONS, PH = 1, 2, 3  # slim-domain analysis IDs, invented
NAMES = {ELEMENTS: "36 Elements", ANIONS: "7 Anions", PH: "pH"}
ACME = 7
TODAY = date(2026, 10, 24)
ND, FIVE = ProcessingTime.NEXT_DAY, ProcessingTime.FIVE_DAYS
_price_ids = count(1)  # unsaved quotes still need distinct price IDs: the engine merges on them


def line(*analysis_ids, prices, bulk_min=0):
    return QuoteLineItem(
        analysis_ids=analysis_ids,
        bulk_price_min=bulk_min,
        prices=tuple(
            QuotePrice(
                id=next(_price_ids),
                processing_time=pt,
                regular_price=Decimal(r),
                bulk_price=None if b is None else Decimal(b),
            )
            for pt, r, b in prices
        ),
    )


def quote(*lines, customers=(), default=False, terms=0, start=date(2026, 1, 1), end=date(2027, 12, 31)):
    return Quote(
        effective_date=start,
        expire_date=end,
        payment_terms=terms,
        is_default=default,
        customer_ids=customers,
        line_items=lines,
        created_by="SUE",
    )


def sample(position, *analysis_ids, pt=ND):
    return BillableSample(position=position, name=f"W-{position:03d}", processing_time=pt, analysis_ids=analysis_ids)


def request(*samples, reported=None, complete=True, tr="TR00042"):
    reported = {(s.position, a) for s in samples for a in s.analysis_ids} if reported is None else reported
    return BillableRequest(
        tr_number=tr,
        complete=complete,
        customer_id=ACME,
        customer_name="Acme Env Labs",
        payment_method=PaymentMethod.PURCHASE_ORDER,
        po_number="PO-1",
        requester_name="Jane Roe",
        requester_email="jane@acme.example",
        date_received=date(2026, 10, 20),
        service_date=date(2026, 10, 23),
        samples=samples,
        reported=frozenset(reported),
        analysis_names=NAMES,
    )
