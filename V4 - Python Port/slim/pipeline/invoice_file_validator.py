from pathlib import Path


def validate_invoice_file(path: str | Path) -> None:
    """Raise ValueError if the path is not a valid invoice submission file."""
    path = Path(path)
    filename = path.name
    _check_extension(path)
    _check_filename_structure(filename)


def _check_extension(path: Path) -> None:
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Invalid file type. Only Excel testing request files allowed.")


def _check_filename_structure(filename: str) -> None:
    if "Results" not in filename:
        raise ValueError("Filename does not contain 'Results'.")
    if "Partial" in filename:
        raise ValueError("File appears to be partial or draft. Cannot invoice.")
    if len(filename) < 6 or not filename[:6].isdigit():
        raise ValueError("File name does not begin with date.")
