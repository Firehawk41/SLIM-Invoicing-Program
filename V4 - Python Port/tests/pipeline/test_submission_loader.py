from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from slim.pipeline.submission_loader import SubmissionLoader, _extract_file_date


# ---------------------------------------------------------------------------
# _extract_file_date
# ---------------------------------------------------------------------------


def test_extract_date_valid():
    assert _extract_file_date("050126 Results.xlsx") == date(2026, 5, 1)


def test_extract_date_end_of_year():
    assert _extract_file_date("123126 Results.xlsx") == date(2026, 12, 31)


def test_extract_date_start_of_year():
    assert _extract_file_date("010126 Results.xlsx") == date(2026, 1, 1)


def test_extract_date_too_short_returns_none():
    assert _extract_file_date("05012.xlsx") is None


def test_extract_date_non_numeric_returns_none():
    assert _extract_file_date("Results050126.xlsx") is None


def test_extract_date_invalid_month_returns_none():
    # month=13 → date() raises ValueError
    assert _extract_file_date("130126 Results.xlsx") is None


def test_extract_date_invalid_day_returns_none():
    # day=99 → date() raises ValueError
    assert _extract_file_date("059926 Results.xlsx") is None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_loader(build_result=None, build_side_effect=None) -> SubmissionLoader:
    svc = MagicMock()
    if build_side_effect is not None:
        svc.build_from_file.side_effect = build_side_effect
    else:
        svc.build_from_file.return_value = build_result or MagicMock()
    return SubmissionLoader(svc)


def _touch(folder: Path, name: str) -> Path:
    path = folder / name
    path.touch()
    return path


# ---------------------------------------------------------------------------
# load_single
# ---------------------------------------------------------------------------


def test_load_single_returns_submission(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    expected = MagicMock()
    loader = _make_loader(build_result=expected)
    result = loader.load_single(tmp_path / "050126 Results.xlsx")
    assert result is expected


def test_load_single_passes_path_to_service(tmp_path):
    path = _touch(tmp_path, "050126 Results.xlsx")
    svc = MagicMock()
    loader = SubmissionLoader(svc)
    try:
        loader.load_single(path)
    except Exception:
        pass
    svc.build_from_file.assert_called_once_with(path)


def test_load_single_invalid_extension_raises(tmp_path):
    _touch(tmp_path, "050126 Results.xls")
    loader = _make_loader()
    with pytest.raises(ValueError, match="file type"):
        loader.load_single(tmp_path / "050126 Results.xls")


def test_load_single_missing_results_raises(tmp_path):
    _touch(tmp_path, "050126 Form.xlsx")
    loader = _make_loader()
    with pytest.raises(ValueError, match="Results"):
        loader.load_single(tmp_path / "050126 Form.xlsx")


def test_load_single_partial_raises(tmp_path):
    _touch(tmp_path, "050126 Results Partial.xlsx")
    loader = _make_loader()
    with pytest.raises(ValueError, match="partial or draft"):
        loader.load_single(tmp_path / "050126 Results Partial.xlsx")


def test_load_single_accepts_string_path(tmp_path):
    path = _touch(tmp_path, "050126 Results.xlsx")
    expected = MagicMock()
    loader = _make_loader(build_result=expected)
    result = loader.load_single(str(path))
    assert result is expected


# ---------------------------------------------------------------------------
# load_by_date_range — date filtering
# ---------------------------------------------------------------------------


def test_load_by_date_range_returns_file_in_range(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    expected = MagicMock()
    loader = _make_loader(build_result=expected)
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == [expected]


def test_load_by_date_range_excludes_before_start(tmp_path):
    _touch(tmp_path, "043026 Results.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


def test_load_by_date_range_excludes_after_end(tmp_path):
    _touch(tmp_path, "060126 Results.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


def test_load_by_date_range_inclusive_start_boundary(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert len(results) == 1


def test_load_by_date_range_inclusive_end_boundary(tmp_path):
    _touch(tmp_path, "053126 Results.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert len(results) == 1


def test_load_by_date_range_multiple_files(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    _touch(tmp_path, "050226 Results.xlsx")
    _touch(tmp_path, "050326 Results.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert len(results) == 3


def test_load_by_date_range_empty_folder(tmp_path):
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


# ---------------------------------------------------------------------------
# load_by_date_range — file validation filtering
# ---------------------------------------------------------------------------


def test_load_by_date_range_skips_invalid_extension(tmp_path):
    _touch(tmp_path, "050126 Results.xls")
    loader = _make_loader()
    # .xls won't be picked up by glob("*.xlsx"), so this is just a sanity check
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


def test_load_by_date_range_skips_missing_results(tmp_path):
    _touch(tmp_path, "050126 Form.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


def test_load_by_date_range_skips_partial(tmp_path):
    _touch(tmp_path, "050126 Results Partial.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


def test_load_by_date_range_skips_non_numeric_prefix(tmp_path):
    _touch(tmp_path, "Results050126.xlsx")
    loader = _make_loader()
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []


# ---------------------------------------------------------------------------
# load_by_date_range — per-file error handling
# ---------------------------------------------------------------------------


def test_load_by_date_range_continues_after_per_file_error(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    _touch(tmp_path, "050226 Results.xlsx")

    call_count = 0

    def side_effect(path):
        nonlocal call_count
        call_count += 1
        if call_count == 1:
            raise ValueError("Simulated parse error")
        return MagicMock()

    loader = _make_loader(build_side_effect=side_effect)
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert len(results) == 1


def test_load_by_date_range_all_errors_returns_empty(tmp_path):
    _touch(tmp_path, "050126 Results.xlsx")
    _touch(tmp_path, "050226 Results.xlsx")
    loader = _make_loader(build_side_effect=RuntimeError("DB error"))
    results = loader.load_by_date_range(tmp_path, date(2026, 5, 1), date(2026, 5, 31))
    assert results == []
