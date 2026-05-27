import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from infrastructure.database import Base
from slim.domain.analysis import AnalysisService
from slim.domain.analysis.analysis_repository import _AnalysisRow, _FormAnalysisRow


@pytest.fixture
def session() -> Session:
    engine = create_engine("sqlite:///:memory:")
    # _AnalysisRow and _FormAnalysisRow are imported above, registering both tables.
    Base.metadata.create_all(engine)
    with Session(engine) as s:
        s.add_all(
            [
                _AnalysisRow(id=1, analysis_name="36 Elements", analysis_description="36 elements (by ICP-MS)"),
                _AnalysisRow(id=11, analysis_name="7 Anions", analysis_description="Anions"),
                _AnalysisRow(id=23, analysis_name="pH", analysis_description=None),
            ]
        )
        s.commit()
        yield s


@pytest.fixture
def service(session: Session) -> AnalysisService:
    return AnalysisService(session)


def test_load_by_id_returns_correct_analysis(service: AnalysisService) -> None:
    a = service.load_analysis(1)
    assert a is not None
    assert a.id == 1
    assert a.name == "36 Elements"
    assert a.description == "36 elements (by ICP-MS)"


def test_load_by_id_returns_none_for_missing(service: AnalysisService) -> None:
    assert service.load_analysis(999) is None


def test_get_by_name_exact_match(service: AnalysisService) -> None:
    a = service.get_by_name("7 Anions")
    assert a is not None
    assert a.id == 11


def test_get_by_name_case_insensitive(service: AnalysisService) -> None:
    a = service.get_by_name("7 anions")
    assert a is not None
    assert a.id == 11


def test_get_by_name_strips_whitespace(service: AnalysisService) -> None:
    a = service.get_by_name("  36 Elements  ")
    assert a is not None
    assert a.id == 1


def test_get_by_name_returns_none_for_missing(service: AnalysisService) -> None:
    assert service.get_by_name("Nonexistent") is None


def test_null_description_coerced_to_empty_string(service: AnalysisService) -> None:
    a = service.load_analysis(23)
    assert a is not None
    assert a.description == ""


def test_all_analyses_returns_all_rows(service: AnalysisService) -> None:
    assert len(service.all_analyses()) == 3


def test_analysis_is_frozen(service: AnalysisService) -> None:
    a = service.load_analysis(1)
    assert a is not None
    with pytest.raises(Exception):
        a.name = "Changed"  # type: ignore[misc]


# --- form_analyses tests ---


def test_form_name_exists_returns_false_for_unmapped(service: AnalysisService) -> None:
    assert service.form_name_exists("36 Elements") is False


def test_get_ids_by_form_name_returns_empty_for_unmapped(service: AnalysisService) -> None:
    assert service.get_ids_by_form_name("36 Elements") == []


def test_add_form_analysis_makes_it_findable(service: AnalysisService) -> None:
    service.add_form_analysis("36 Elements", 1)

    assert service.form_name_exists("36 Elements") is True
    assert service.get_ids_by_form_name("36 Elements") == [1]


def test_add_form_analysis_case_insensitive_lookup(service: AnalysisService) -> None:
    service.add_form_analysis("36 Elements", 1)

    assert service.get_ids_by_form_name("36 ELEMENTS") == [1]
    assert service.get_ids_by_form_name("36 elements") == [1]


def test_add_form_analysis_is_idempotent(service: AnalysisService) -> None:
    service.add_form_analysis("36 Elements", 1)
    service.add_form_analysis("36 Elements", 1)  # must not raise or duplicate

    assert service.get_ids_by_form_name("36 Elements") == [1]


def test_add_form_analysis_multiple_analyses_per_form_name(service: AnalysisService) -> None:
    # A bundle form name maps to more than one analysis ID
    service.add_form_analysis("Bundle A", 1)
    service.add_form_analysis("Bundle A", 11)

    ids = service.get_ids_by_form_name("Bundle A")
    assert set(ids) == {1, 11}


def test_add_form_analysis_survives_db_roundtrip(session: Session) -> None:
    svc1 = AnalysisService(session)
    svc1.add_form_analysis("36 Elements", 1)

    # Fresh service rebuilds cache from DB — proves the row was committed
    svc2 = AnalysisService(session)
    assert svc2.form_name_exists("36 Elements") is True
    assert svc2.get_ids_by_form_name("36 Elements") == [1]


def test_cache_loads_form_analyses_seeded_before_startup(session: Session) -> None:
    # Seed form_analyses rows directly (simulates rows already in the DB)
    session.add_all([
        _FormAnalysisRow(form_name="36 Elements", analysis_id=1),
        _FormAnalysisRow(form_name="36 Elements", analysis_id=11),
    ])
    session.commit()

    svc = AnalysisService(session)
    assert svc.form_name_exists("36 Elements") is True
    assert set(svc.get_ids_by_form_name("36 Elements")) == {1, 11}
