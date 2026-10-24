"""What the host hands invoicing: one TR as it was reported (D-83).

The host builds it from its TR and slim-data's report issues; invoicing never reads either. The rule
that matters lives here: bill each ordered analysis of a sample that a report showed, once.
"""

from datetime import date

from pydantic import BaseModel, ConfigDict, Field

from slim_invoicing.domain.enums import PaymentMethod, ProcessingTime


class _Frozen(BaseModel):
    model_config = ConfigDict(frozen=True)


class BillableSample(_Frozen):
    position: int = Field(ge=1)
    name: str = Field(min_length=1)
    processing_time: ProcessingTime
    analysis_ids: tuple[int, ...]  # as ordered


class BillableRequest(_Frozen):
    tr_number: str = Field(min_length=1)
    complete: bool  # the TR's report is complete; only then is it invoiced
    customer_id: int
    customer_name: str = Field(min_length=1)
    payment_method: PaymentMethod
    po_number: str = ""
    requester_name: str = ""
    requester_email: str = ""
    date_received: date
    service_date: date
    samples: tuple[BillableSample, ...]
    reported: frozenset[tuple[int, int]]  # (sample position, analysis ID) a report issue showed
    analysis_names: dict[int, str]  # for the line items' text

    def billable_samples(self) -> tuple[BillableSample, ...]:
        """Each sample with only the ordered analyses that were reported; samples with none drop out.

        `reported` is a set, so an analysis shown on an amended report as well as the original counts
        once, and reruns or failed attempts (never reported) never count at all.
        """
        out = []
        for sample in self.samples:
            billed = tuple(a for a in sample.analysis_ids if (sample.position, a) in self.reported)
            if billed:
                out.append(sample.model_copy(update={"analysis_ids": billed}))
        return tuple(out)
