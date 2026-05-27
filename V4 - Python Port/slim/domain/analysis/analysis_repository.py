from sqlalchemy import ForeignKey, String
from sqlalchemy.orm import Mapped, Session, mapped_column

from infrastructure.database import Base
from slim.domain.analysis.analysis import Analysis


class _AnalysisRow(Base):
    __tablename__ = "analyses"

    id: Mapped[int] = mapped_column("ID", primary_key=True)
    analysis_name: Mapped[str] = mapped_column(String)
    analysis_description: Mapped[str | None] = mapped_column(String, nullable=True)


class _FormAnalysisRow(Base):
    __tablename__ = "form_analyses"

    form_name: Mapped[str] = mapped_column(String, primary_key=True)
    analysis_id: Mapped[int] = mapped_column(ForeignKey("analyses.ID"), primary_key=True)


class AnalysisRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def select_all(self) -> list[Analysis]:
        """Load every analysis row ordered by name. Called only by AnalysisCache.build."""
        rows = (
            self._session.query(_AnalysisRow)
            .order_by(_AnalysisRow.analysis_name)
            .all()
        )
        return [
            Analysis(
                id=r.id,
                name=r.analysis_name,
                description=r.analysis_description or "",
            )
            for r in rows
        ]

    def select_all_form_analyses(self) -> list[tuple[str, int]]:
        """Load all form-name → analysis_id mappings. Called only by AnalysisCache."""
        rows = self._session.query(_FormAnalysisRow).all()
        return [(r.form_name, r.analysis_id) for r in rows]

    def insert_form_analysis(self, form_name: str, analysis_id: int) -> None:
        """Persist a new form-name → analysis_id mapping."""
        self._session.add(_FormAnalysisRow(form_name=form_name, analysis_id=analysis_id))
        self._session.commit()
