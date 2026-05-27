from datetime import datetime

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.chemical.chemical_repository import _ChemicalRow, _FormChemicalRow, _KEDElementRow
from slim.domain.chemical.chemical_service import ChemicalService

SEED_DATE = datetime(2024, 1, 15, 10, 30, 0)


@pytest.fixture
def session() -> Session:
    engine = create_engine("sqlite:///:memory:")
    Base.metadata.create_all(engine)
    with Session(engine) as s:
        s.add_all(
            [
                _ChemicalRow(
                    id=1,
                    chemical_name="Hydrochloric Acid",
                    metals_prep="Dilute and Shoot",
                    silicon_prep="N/A",
                    ions_prep="Dilute and Shoot",
                    database_entry_date=SEED_DATE,
                ),
                _ChemicalRow(
                    id=2,
                    chemical_name="Nitric Acid",
                    metals_prep="Evaporation",
                    silicon_prep="Evaporation",
                    ions_prep="Dilute and Shoot",
                    database_entry_date=None,
                ),
                _KEDElementRow(chemical_id=1, element_id=10),
                _KEDElementRow(chemical_id=1, element_id=20),
                _KEDElementRow(chemical_id=1, element_id=30),
            ]
        )
        s.commit()
        yield s


@pytest.fixture
def service(session: Session) -> ChemicalService:
    return ChemicalService(session)


# --- Read tests ---


def test_load_by_id_returns_correct_chemical(service: ChemicalService) -> None:
    c = service.load_chemical(1)
    assert c is not None
    assert c.id == 1
    assert c.name == "Hydrochloric Acid"
    assert c.metals_prep == "Dilute and Shoot"
    assert c.silicon_prep == "N/A"
    assert c.ions_prep == "Dilute and Shoot"
    assert c.entry_date == SEED_DATE


def test_load_by_id_returns_none_for_missing(service: ChemicalService) -> None:
    assert service.load_chemical(999) is None


def test_get_by_name_exact_match(service: ChemicalService) -> None:
    c = service.get_by_name("Nitric Acid")
    assert c is not None
    assert c.id == 2


def test_get_by_name_case_insensitive(service: ChemicalService) -> None:
    c = service.get_by_name("nitric acid")
    assert c is not None
    assert c.id == 2


def test_get_by_name_returns_none_for_missing(service: ChemicalService) -> None:
    assert service.get_by_name("Unobtanium Chloride") is None


def test_ked_element_ids_loaded_correctly(service: ChemicalService) -> None:
    c = service.load_chemical(1)
    assert c is not None
    assert c.ked_element_ids == frozenset({10, 20, 30})


def test_chemical_with_no_ked_elements(service: ChemicalService) -> None:
    c = service.load_chemical(2)
    assert c is not None
    assert c.ked_element_ids == frozenset()


def test_null_entry_date_allowed(service: ChemicalService) -> None:
    c = service.load_chemical(2)
    assert c is not None
    assert c.entry_date is None


def test_all_chemicals_returns_all_rows(service: ChemicalService) -> None:
    assert len(service.all_chemicals()) == 2


def test_chemical_is_frozen(service: ChemicalService) -> None:
    c = service.load_chemical(1)
    assert c is not None
    with pytest.raises(Exception):
        c.name = "Changed"  # type: ignore[misc]


# --- Write tests ---


def test_create_chemical_returns_persisted_entity(service: ChemicalService) -> None:
    c = service.create_chemical(
        name="Sulfuric Acid",
        metals_prep="Evaporation",
        silicon_prep="N/A",
        ions_prep="Dilute and Shoot",
    )
    assert c.id > 0
    assert c.name == "Sulfuric Acid"
    assert c.entry_date is not None


def test_create_chemical_is_findable_via_cache(service: ChemicalService) -> None:
    service.create_chemical(
        name="Sulfuric Acid",
        metals_prep="Evaporation",
        silicon_prep="N/A",
        ions_prep="Dilute and Shoot",
    )
    assert service.get_by_name("Sulfuric Acid") is not None


def test_create_chemical_is_persisted_to_db(session: Session) -> None:
    service1 = ChemicalService(session)
    service1.create_chemical(
        name="Acetic Acid",
        metals_prep="Dilute and Shoot",
        silicon_prep="N/A",
        ions_prep="Dilute and Shoot",
    )
    # Fresh service rebuilds cache from DB — proves the row was committed
    service2 = ChemicalService(session)
    assert service2.get_by_name("Acetic Acid") is not None


def test_create_chemical_with_ked_elements(service: ChemicalService) -> None:
    c = service.create_chemical(
        name="Phosphoric Acid",
        metals_prep="Dilute and Shoot",
        silicon_prep="Dilute and Shoot",
        ions_prep="N/A",
        ked_element_ids=frozenset({5, 15}),
    )
    assert c.ked_element_ids == frozenset({5, 15})


def test_create_chemical_ked_elements_survive_db_roundtrip(session: Session) -> None:
    service1 = ChemicalService(session)
    c = service1.create_chemical(
        name="Phosphoric Acid",
        metals_prep="Dilute and Shoot",
        silicon_prep="Dilute and Shoot",
        ions_prep="N/A",
        ked_element_ids=frozenset({5, 15}),
    )
    service2 = ChemicalService(session)
    reloaded = service2.load_chemical(c.id)
    assert reloaded is not None
    assert reloaded.ked_element_ids == frozenset({5, 15})


def test_create_chemical_duplicate_name_raises(service: ChemicalService) -> None:
    with pytest.raises(ValueError, match="already exists"):
        service.create_chemical(
            name="Hydrochloric Acid",
            metals_prep="Evaporation",
            silicon_prep="N/A",
            ions_prep="N/A",
        )


def test_create_chemical_blank_name_raises(service: ChemicalService) -> None:
    with pytest.raises(ValueError, match="cannot be blank"):
        service.create_chemical(
            name="  ",
            metals_prep="Evaporation",
            silicon_prep="N/A",
            ions_prep="N/A",
        )


def test_create_chemical_invalid_metals_prep_raises(service: ChemicalService) -> None:
    with pytest.raises(ValueError, match="metals prep"):
        service.create_chemical(
            name="Acetic Acid",
            metals_prep="Deep Fry",
            silicon_prep="N/A",
            ions_prep="N/A",
        )


def test_create_chemical_invalid_silicon_prep_raises(service: ChemicalService) -> None:
    with pytest.raises(ValueError, match="silicon prep"):
        service.create_chemical(
            name="Acetic Acid",
            metals_prep="N/A",
            silicon_prep="Boil",
            ions_prep="N/A",
        )


def test_create_chemical_invalid_ions_prep_raises(service: ChemicalService) -> None:
    with pytest.raises(ValueError, match="ions prep"):
        service.create_chemical(
            name="Acetic Acid",
            metals_prep="N/A",
            silicon_prep="N/A",
            ions_prep="Shake and Bake",
        )


# --- form_chemicals tests ---


def test_form_name_exists_returns_false_for_unmapped(service: ChemicalService) -> None:
    assert service.form_name_exists("HCl") is False


def test_get_id_by_form_name_returns_zero_for_unmapped(service: ChemicalService) -> None:
    assert service.get_id_by_form_name("HCl") == 0


def test_get_by_form_name_returns_none_for_unmapped(service: ChemicalService) -> None:
    assert service.get_by_form_name("HCl") is None


def test_add_form_name_makes_it_findable(service: ChemicalService) -> None:
    service.add_form_name("HCl", 1)

    assert service.form_name_exists("HCl") is True
    assert service.get_id_by_form_name("HCl") == 1
    assert service.get_by_form_name("HCl") is not None
    assert service.get_by_form_name("HCl").id == 1  # type: ignore[union-attr]


def test_add_form_name_case_insensitive_lookup(service: ChemicalService) -> None:
    service.add_form_name("HCl", 1)

    assert service.get_id_by_form_name("hcl") == 1
    assert service.get_id_by_form_name("HCL") == 1


def test_add_form_name_is_idempotent(service: ChemicalService) -> None:
    service.add_form_name("HCl", 1)
    service.add_form_name("HCl", 1)  # must not raise or duplicate

    assert service.get_id_by_form_name("HCl") == 1


def test_add_form_name_survives_db_roundtrip(session: Session) -> None:
    svc1 = ChemicalService(session)
    svc1.add_form_name("HCl", 1)

    # Fresh service rebuilds cache from DB — proves the row was committed
    svc2 = ChemicalService(session)
    assert svc2.form_name_exists("HCl") is True
    assert svc2.get_id_by_form_name("HCl") == 1


def test_cache_loads_form_names_seeded_before_startup(session: Session) -> None:
    # Seed a form_chemicals row directly (simulates a row that already exists in the DB)
    session.add(_FormChemicalRow(form_name="Hydrochloric Acid", chemical_id=1))
    session.commit()

    svc = ChemicalService(session)
    assert svc.form_name_exists("Hydrochloric Acid") is True
    assert svc.get_id_by_form_name("Hydrochloric Acid") == 1
    assert svc.get_by_form_name("Hydrochloric Acid") is not None
