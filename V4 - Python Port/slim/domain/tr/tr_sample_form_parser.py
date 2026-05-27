from typing import Any, NamedTuple

from slim.domain.analysis.analysis_service import AnalysisService
from slim.domain.chemical.chemical_service import ChemicalService
from slim.domain.element.element_service import ElementService
from slim.domain.tr.enums import ProcessingTime, RequestType
from slim.domain.tr.tr_sample import TRSample


class SampleRangeBounds(NamedTuple):
    first_row: int
    last_row: int
    first_col: int  # always col 2 (sample-ID column)
    last_col: int


class TRSampleFormParser:
    def __init__(
        self,
        chemical_service: ChemicalService,
        analysis_service: AnalysisService,
        element_service: ElementService,
        input_resolver: Any,  # TRFormInputResolver or test stub
    ) -> None:
        self._chemical_svc = chemical_service
        self._analysis_svc = analysis_service
        self._element_svc = element_service
        self._resolver = input_resolver

    def build_from_row(
        self,
        ws: Any,  # openpyxl Worksheet
        row: int,
        sample_range: SampleRangeBounds,
        request_type: RequestType,
    ) -> TRSample:
        form_chem_name = _get_form_chemical_name(ws, row, request_type)
        chemical = self._resolver.resolve_chemical(form_chem_name)

        processing_time = _get_processing_time(ws, row, request_type)
        additional_notes = _cell_str(ws, row, _notes_col(request_type))
        requested_time = _cell_str(ws, row, _requested_time_col(request_type))
        sample_name = _cell_str(ws, row, 2)

        analysis_ids = self._collect_analysis_ids(ws, row, sample_range, request_type)
        additional_element_ids = self._collect_additional_element_ids(ws, row, request_type)

        return TRSample(
            id=0,
            sample_name=sample_name,
            form_chemical_name=form_chem_name,
            processing_time=processing_time,
            additional_notes=additional_notes,
            requested_time=requested_time,
            chemical_id=chemical.id,
            analysis_ids=tuple(analysis_ids),
            additional_element_ids=tuple(additional_element_ids),
        )

    def _collect_analysis_ids(
        self,
        ws: Any,
        row: int,
        sample_range: SampleRangeBounds,
        request_type: RequestType,
    ) -> list[int]:
        start_col = _analysis_start_col(request_type)
        seen: set[int] = set()
        result: list[int] = []

        for col in range(start_col, sample_range.last_col + 1):
            val = ws.cell(row=row, column=col).value
            if val is None or not str(val).strip():
                continue
            try:
                ids = self._analysis_svc.get_ids_by_form_name(str(val).strip())
                for aid in ids:
                    if aid not in seen:
                        seen.add(aid)
                        result.append(aid)
            except NotImplementedError:
                pass  # form_analyses not yet ported; analysis lookup will work once it is

        # Check if the additional-elements column header itself maps to an analysis.
        add_el_col = _additional_element_col(request_type)
        if add_el_col > 0:
            header_val = ws.cell(row=sample_range.first_row - 1, column=add_el_col).value
            if header_val and str(header_val).strip():
                analysis = self._analysis_svc.get_by_name(str(header_val).strip())
                if analysis is not None and analysis.id not in seen:
                    seen.add(analysis.id)
                    result.append(analysis.id)

        return result

    def _collect_additional_element_ids(
        self,
        ws: Any,
        row: int,
        request_type: RequestType,
    ) -> list[int]:
        add_el_col = _additional_element_col(request_type)
        if add_el_col == 0:
            return []
        cell_val = str(ws.cell(row=row, column=add_el_col).value or "").strip()
        if not cell_val:
            return []
        seen: set[int] = set()
        result: list[int] = []
        for symbol in _parse_element_symbols(cell_val):
            element = self._element_svc.get_by_symbol(symbol)
            if element is not None and element.id not in seen:
                seen.add(element.id)
                result.append(element.id)
        return result


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def _cell_str(ws: Any, row: int, col: int) -> str:
    if col == 0:
        return ""
    return str(ws.cell(row=row, column=col).value or "")


def _get_form_chemical_name(ws: Any, row: int, request_type: RequestType) -> str:
    if request_type == RequestType.CHEMICAL:
        return _cell_str(ws, row, 3)
    if request_type == RequestType.WATER:
        return "Water"
    return _cell_str(ws, row, 4)  # Wafer


def _get_processing_time(ws: Any, row: int, request_type: RequestType) -> ProcessingTime:
    col_map = {RequestType.CHEMICAL: 11, RequestType.WATER: 14, RequestType.WAFER: 9}
    raw = _cell_str(ws, row, col_map[request_type])
    return ProcessingTime.from_form_string(raw)


def _notes_col(request_type: RequestType) -> int:
    return {RequestType.CHEMICAL: 14, RequestType.WATER: 17, RequestType.WAFER: 10}[request_type]


def _requested_time_col(request_type: RequestType) -> int:
    return {RequestType.CHEMICAL: 13, RequestType.WATER: 16, RequestType.WAFER: 0}[request_type]


def _analysis_start_col(request_type: RequestType) -> int:
    return {RequestType.CHEMICAL: 4, RequestType.WATER: 4, RequestType.WAFER: 5}[request_type]


def _additional_element_col(request_type: RequestType) -> int:
    if request_type in (RequestType.CHEMICAL, RequestType.WATER):
        return 5
    if request_type == RequestType.WAFER:
        return 6
    return 0


def _parse_element_symbols(element_string: str) -> list[str]:
    """Split a delimited element symbol string (;  .  space  newline) into tokens."""
    s = element_string
    for sep in (";", ".", " ", "\r\n", "\r", "\n"):
        s = s.replace(sep, ",")
    return [t.strip() for t in s.split(",") if t.strip()]
