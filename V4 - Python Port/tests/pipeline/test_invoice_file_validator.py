from pathlib import Path

import pytest

from slim.pipeline.invoice_file_validator import validate_invoice_file


def test_valid_file_passes():
    validate_invoice_file(Path("050126 Results.xlsx"))


def test_valid_file_string_path():
    validate_invoice_file("050126 Results.xlsx")


def test_wrong_extension_xls_fails():
    with pytest.raises(ValueError, match="file type"):
        validate_invoice_file(Path("050126 Results.xls"))


def test_wrong_extension_csv_fails():
    with pytest.raises(ValueError, match="file type"):
        validate_invoice_file(Path("050126 Results.csv"))


def test_no_extension_fails():
    with pytest.raises(ValueError, match="file type"):
        validate_invoice_file(Path("050126 Results"))


def test_extension_case_insensitive():
    # .XLSX uppercase should pass
    validate_invoice_file(Path("050126 Results.XLSX"))


def test_missing_results_keyword_fails():
    with pytest.raises(ValueError, match="Results"):
        validate_invoice_file(Path("050126 Form.xlsx"))


def test_partial_in_name_fails():
    with pytest.raises(ValueError, match="partial or draft"):
        validate_invoice_file(Path("050126 Results Partial.xlsx"))


def test_partial_substring_case_sensitive():
    # "Partial" is case-sensitive match in VBA — same in Python
    validate_invoice_file(Path("050126 Results partial.xlsx"))


def test_non_numeric_prefix_fails():
    with pytest.raises(ValueError, match="date"):
        validate_invoice_file(Path("Results050126.xlsx"))


def test_short_filename_fails():
    with pytest.raises(ValueError, match="date"):
        validate_invoice_file(Path("05012 Results.xlsx"))


def test_six_digit_prefix_required():
    validate_invoice_file(Path("050126 Results Detail.xlsx"))


def test_results_and_partial_partial_wins():
    # "Partial" check runs after "Results" check — partial rejection takes priority
    with pytest.raises(ValueError, match="partial or draft"):
        validate_invoice_file(Path("050126 Results Partial Draft.xlsx"))
