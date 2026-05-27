import logging
from datetime import date
from pathlib import Path

from slim.domain.tr.tr_submission import TRSubmission
from slim.domain.tr.tr_submission_service import TRSubmissionService
from slim.pipeline.invoice_file_validator import validate_invoice_file

_logger = logging.getLogger(__name__)


def _extract_file_date(filename: str) -> date | None:
    """Extract date from a MMDDYY filename prefix. Returns None if not parseable."""
    if len(filename) < 6 or not filename[:6].isdigit():
        return None
    try:
        month = int(filename[:2])
        day = int(filename[2:4])
        year = int(filename[4:6]) + 2000
        return date(year, month, day)
    except ValueError:
        return None


class SubmissionLoader:
    """Loads TRSubmission objects from .xlsx files on disk."""

    def __init__(self, tr_svc: TRSubmissionService) -> None:
        self._tr_svc = tr_svc

    def load_single(self, file_path: str | Path) -> TRSubmission:
        """Validate and build one TRSubmission from a single file. Raises on invalid file."""
        path = Path(file_path)
        validate_invoice_file(path)
        return self._tr_svc.build_from_file(path)

    def load_by_date_range(
        self,
        folder_path: str | Path,
        start_date: date,
        end_date: date,
    ) -> list[TRSubmission]:
        """Scan folder for eligible .xlsx files by MMDDYY date prefix, build submissions.

        Per-file errors are logged and skipped; the batch continues.
        """
        folder = Path(folder_path)
        results: list[TRSubmission] = []
        scanned = eligible = created = skipped = errors = 0

        _logger.info("Scanning folder: %s", folder)

        for path in sorted(folder.glob("*.xlsx")):
            scanned += 1
            _logger.debug("Evaluating file: %s", path.name)

            try:
                validate_invoice_file(path)
            except ValueError as e:
                _logger.warning("Invalid file skipped: %s | %s", path.name, e)
                skipped += 1
                continue

            file_date = _extract_file_date(path.name)
            if file_date is None:
                _logger.warning("Could not extract date from filename: %s", path.name)
                skipped += 1
                continue

            if not start_date <= file_date <= end_date:
                skipped += 1
                continue

            eligible += 1
            _logger.debug("File passed eligibility checks: %s", path.name)

            try:
                submission = self._tr_svc.build_from_file(path)
                results.append(submission)
                created += 1
                _logger.info("Submission added: %s", path.name)
            except Exception as e:
                errors += 1
                _logger.error("Error building submission from %s: %s", path.name, e)

        _logger.info(
            "Batch summary | Files scanned: %d | Eligible: %d | "
            "Submissions created: %d | Skipped: %d | Errors: %d",
            scanned,
            eligible,
            created,
            skipped,
            errors,
        )

        return results
