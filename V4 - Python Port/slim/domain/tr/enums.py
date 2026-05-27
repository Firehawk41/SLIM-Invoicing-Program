from enum import IntEnum


class RequestType(IntEnum):
    CHEMICAL = 1
    WATER = 2
    WAFER = 3


class ProcessingTime(IntEnum):
    EXTENDED_TIME = 1
    NEXT_DAY = 2
    TIME_LIMITED = 3
    SAME_DAY_RUSH = 4
    CALL_IN_RUSH = 5
    TWO_DAYS = 6
    THREE_DAYS = 7
    FIVE_DAYS = 8

    @property
    def days(self) -> int:
        return _PT_DAYS[self]

    @property
    def label(self) -> str:
        return _PT_LABELS[self]

    @classmethod
    def from_form_string(cls, value: str) -> "ProcessingTime":
        key = value.lower().strip()
        if key not in _PT_FROM_FORM:
            raise ValueError(f"Unrecognised processing time: {value!r}")
        return _PT_FROM_FORM[key]


_PT_DAYS: dict[ProcessingTime, int] = {
    ProcessingTime.SAME_DAY_RUSH: 0,
    ProcessingTime.CALL_IN_RUSH: 0,
    ProcessingTime.NEXT_DAY: 1,
    ProcessingTime.TIME_LIMITED: 1,
    ProcessingTime.TWO_DAYS: 2,
    ProcessingTime.EXTENDED_TIME: 3,
    ProcessingTime.THREE_DAYS: 3,
    ProcessingTime.FIVE_DAYS: 5,
}

_PT_LABELS: dict[ProcessingTime, str] = {
    ProcessingTime.EXTENDED_TIME: "Extended Time",
    ProcessingTime.NEXT_DAY: "Next Day",
    ProcessingTime.TIME_LIMITED: "Time Limited",
    ProcessingTime.SAME_DAY_RUSH: "Same Day Rush",
    ProcessingTime.CALL_IN_RUSH: "Call In Rush",
    ProcessingTime.TWO_DAYS: "Two Days",
    ProcessingTime.THREE_DAYS: "Three Days",
    ProcessingTime.FIVE_DAYS: "Five Days",
}

_PT_FROM_FORM: dict[str, ProcessingTime] = {
    "extended time": ProcessingTime.EXTENDED_TIME,
    "next day": ProcessingTime.NEXT_DAY,
    "next day rush": ProcessingTime.NEXT_DAY,
    "time limited": ProcessingTime.TIME_LIMITED,
    "timelimited": ProcessingTime.TIME_LIMITED,
    "next day time limited": ProcessingTime.TIME_LIMITED,
    "same day rush": ProcessingTime.SAME_DAY_RUSH,
    "samedayrush": ProcessingTime.SAME_DAY_RUSH,
    "call in rush": ProcessingTime.CALL_IN_RUSH,
    "callinrush": ProcessingTime.CALL_IN_RUSH,
    "two days": ProcessingTime.TWO_DAYS,
    "2 days": ProcessingTime.TWO_DAYS,
    "2days": ProcessingTime.TWO_DAYS,
    "three days": ProcessingTime.THREE_DAYS,
    "3 days": ProcessingTime.THREE_DAYS,
    "3days": ProcessingTime.THREE_DAYS,
    "up to 3 working days": ProcessingTime.THREE_DAYS,
    "five days": ProcessingTime.FIVE_DAYS,
    "5 days": ProcessingTime.FIVE_DAYS,
    "5days": ProcessingTime.FIVE_DAYS,
}
