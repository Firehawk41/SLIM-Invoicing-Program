from pydantic import BaseModel, ConfigDict

from slim.domain.tr.enums import ProcessingTime


class TRSample(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int = 0
    sample_name: str
    form_chemical_name: str
    processing_time: ProcessingTime
    additional_notes: str
    requested_time: str  # empty string for Wafer (column not present)
    chemical_id: int
    analysis_ids: tuple[int, ...] = ()
    additional_element_ids: tuple[int, ...] = ()

    @property
    def processing_time_str(self) -> str:
        return self.processing_time.label
