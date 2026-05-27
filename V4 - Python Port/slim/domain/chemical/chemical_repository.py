from datetime import datetime

from sqlalchemy import ForeignKey, String
from sqlalchemy.orm import Mapped, Session, mapped_column

from infrastructure.database import Base
from slim.domain.chemical.chemical import Chemical


class _ChemicalRow(Base):
    __tablename__ = "chemicals"

    id: Mapped[int] = mapped_column("ID", primary_key=True)
    chemical_name: Mapped[str] = mapped_column(String)
    metals_prep: Mapped[str] = mapped_column(String)
    silicon_prep: Mapped[str] = mapped_column(String)
    ions_prep: Mapped[str] = mapped_column(String)
    database_entry_date: Mapped[datetime | None] = mapped_column(nullable=True)


class _KEDElementRow(Base):
    __tablename__ = "ked_elements"

    chemical_id: Mapped[int] = mapped_column(ForeignKey("chemicals.ID"), primary_key=True)
    element_id: Mapped[int] = mapped_column(primary_key=True)


class _FormChemicalRow(Base):
    __tablename__ = "form_chemicals"

    form_name: Mapped[str] = mapped_column(String, primary_key=True)
    chemical_id: Mapped[int] = mapped_column(ForeignKey("chemicals.ID"))


class ChemicalRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def select_all(self) -> list[Chemical]:
        """Load all chemicals and KED element IDs in two queries. Called only by ChemicalCache."""
        chem_rows = (
            self._session.query(_ChemicalRow)
            .order_by(_ChemicalRow.chemical_name)
            .all()
        )
        ked_rows = self._session.query(_KEDElementRow).all()

        ked_map: dict[int, set[int]] = {}
        for r in ked_rows:
            ked_map.setdefault(r.chemical_id, set()).add(r.element_id)

        return [
            Chemical(
                id=r.id,
                name=r.chemical_name,
                metals_prep=r.metals_prep,
                silicon_prep=r.silicon_prep,
                ions_prep=r.ions_prep,
                entry_date=r.database_entry_date,
                ked_element_ids=frozenset(ked_map.get(r.id, set())),
            )
            for r in chem_rows
        ]

    def insert(
        self,
        name: str,
        metals_prep: str,
        silicon_prep: str,
        ions_prep: str,
        ked_element_ids: frozenset[int],
    ) -> Chemical:
        """Insert a new chemical + its KED elements in one transaction. Returns the persisted Chemical."""
        row = _ChemicalRow(
            chemical_name=name,
            metals_prep=metals_prep,
            silicon_prep=silicon_prep,
            ions_prep=ions_prep,
            database_entry_date=datetime.now(),
        )
        self._session.add(row)
        self._session.flush()  # materialise the auto-assigned ID before adding FK children

        for element_id in ked_element_ids:
            self._session.add(_KEDElementRow(chemical_id=row.id, element_id=element_id))

        self._session.commit()
        return Chemical(
            id=row.id,
            name=name,
            metals_prep=metals_prep,
            silicon_prep=silicon_prep,
            ions_prep=ions_prep,
            entry_date=row.database_entry_date,
            ked_element_ids=ked_element_ids,
        )

    def select_all_form_names(self) -> list[tuple[str, int]]:
        """Load all form-name → chemical_id mappings. Called only by ChemicalCache."""
        rows = self._session.query(_FormChemicalRow).all()
        return [(r.form_name, r.chemical_id) for r in rows]

    def insert_form_name(self, form_name: str, chemical_id: int) -> None:
        """Persist a new form-name → chemical_id mapping."""
        self._session.add(_FormChemicalRow(form_name=form_name, chemical_id=chemical_id))
        self._session.commit()
