"""The host's codes, by value: V5 never imports the host, so these mirror its enums (ints in the DB)."""

from enum import IntEnum


class ProcessingTime(IntEnum):
    EXTENDED_TIME = 1
    NEXT_DAY = 2
    TIME_LIMITED = 3
    SAME_DAY_RUSH = 4
    CALL_IN_RUSH = 5
    TWO_DAYS = 6
    THREE_DAYS = 7
    FIVE_DAYS = 8
    NEXT_DAY_RUSH = 9

    @property
    def label(self) -> str:
        return _LABELS[self]


_LABELS = {
    ProcessingTime.EXTENDED_TIME: "Extended Time",
    ProcessingTime.NEXT_DAY: "Next Day",
    ProcessingTime.TIME_LIMITED: "Next Day Time Limited",
    ProcessingTime.SAME_DAY_RUSH: "Same Day Rush",
    ProcessingTime.CALL_IN_RUSH: "Call In Rush",
    ProcessingTime.TWO_DAYS: "Two Days",
    ProcessingTime.THREE_DAYS: "Three Days",
    ProcessingTime.FIVE_DAYS: "Five Days",
    ProcessingTime.NEXT_DAY_RUSH: "Next Day Rush",
}


class PaymentMethod(IntEnum):
    PURCHASE_ORDER = 1
    CREDIT_CARD = 2  # the lab calls the customer; no card data is ever stored
