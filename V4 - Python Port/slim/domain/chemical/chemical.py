from datetime import datetime

from pydantic import BaseModel, ConfigDict


class Chemical(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    name: str
    metals_prep: str
    silicon_prep: str
    ions_prep: str
    entry_date: datetime | None = None
    ked_element_ids: frozenset[int] = frozenset()
