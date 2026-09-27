from __future__ import annotations

import logging
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Callable

from .errors import DataValidationError, OfficeAutomationError, WorkbookCommitError
from .mkr_parser import extract_mkr_numbers, to_excel_sheet_name
from .office import _load_com, assert_workbook_available


TEMPLATE_SHEET = "MKR_TEMPLATE"
EXPECTED_HEADERS = {
    0: "Code",
    1: "Name",
    2: "Order QTY",
    3: "Shipped QTY",
    4: "Shortage QTY",
    7: "Back Order NO.",
    8: "Code",
    9: "Name",
    10: "Order QTY",
    11: "Shipped QTY",
    12: "Shortage QTY",
}


def _sheet_names(workbook) -> set[str]:
    return {
        str(workbook.Worksheets.Item(index).Name)
        for index in range(1, int(workbook.Worksheets.Count) + 1)
    }


def _clean(value: object) -> str:
    return "" if value is None else str(value).strip()


def _as_matrix(value) -> list[list[object]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _merged_ranges(sheet) -> tuple[str, ...]:
    merged: set[str] = set()
    try:
        areas = sheet.UsedRange.MergeAreas
        for index in range(1, int(areas.Count) + 1):
            merged.add(str(areas.Item(index).Address))
    except Exception:
        pass
    return tuple(sorted(merged))


def _structure_fingerprint(sheet) -> tuple:
    used = sheet.UsedRange
    headers = tuple(_clean(value) for value in sheet.Range("A9:M9").Value2[0])
    return (
        str(used.Address),
        int(used.Rows.Count),
        int(used.Columns.Count),
        _merged_ranges(sheet),
        headers,
    )


def _validate_headers(sheet, label: str) -> None:
    values = tuple(_clean(value) for value in sheet.Range("A9:M9").Value2[0])
    for index, expected in EXPECTED_HEADERS.items():
        if index >= len(values) or values[index] != expected:
            actual = values[index] if index < len(values) else ""
            raise DataValidationError(
                f"{label}의 {chr(65 + index)}9 헤더가 올바르지 않습니다: "
                f"예상={expected} / 실제={actual or '빈 셀'}"
            )


def _validate_no_external_formulas(sheet) -> None:
    formulas = _as_matrix(sheet.UsedRange.Formula)
    for row in formulas:
        for value in row:
            if isinstance(value, str) and value.startswith("=") and "[" in value and "]" in value:
                raise DataValidationError(
                    "MKR_TEMPLATE에 외부 통합문서를 참조하는 수식이 있습니다: " + value[:160]
                )


def _validate_existing_sheet(sheet, mkr_number: str, template_fingerprint: tuple | None) -> None:
    current = _clean(sheet.Range("D3").Value2).upper().replace(" ", "")
    expected = mkr_number.upper().replace(" ", "")
    if current != expected:
        raise DataValidationError(
            f"기존 시트 {sheet.Name}의 D3 값이 대상 MKR과 다릅니다: "
            f"예상={mkr_number} / 실제={current or '빈 셀'}"
        )
    _validate_headers(sheet, str(sheet.Name))
    if template_fingerprint is not None and _structure_fingerprint(sheet) != template_fingerprint:
        raise DataValidationError(f"기존 시트 {sheet.Name}의 구조가 외부 템플릿과 다릅니다.")


def _disk_sheet_names(path: Path) -> set[str] | None:
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return None
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    return {node.attrib["name"] for node in root.findall("x:sheets/x:sheet", namespace)}


def _backup_path(workbook_path: Path, backup_dir: Path | None) -> Path:
    destination = backup_dir or workbook_path.parent / "Backup"
    destination.mkdir(parents=True, exist_ok=True)
    return destination / (
        f"{workbook_path.stem}_before_sheet_create_"
        f"{datetime.now():%Y%m%d_%H%M%S_%f}{workbook_path.suffix}"
    )


def _normalize_numbers(values: list[str]) -> list[tuple[str, str]]:
    normalized: list[tuple[str, str]] = []
    seen: set[str] = set()
    for raw in values:
        parsed = extract_mkr_numbers(raw)
        if len(parsed) != 1:
            raise DataValidationError(f"올바른 MKR 번호가 아닙니다: {raw}")
        number = parsed[0]
        if number in seen:
            continue
        seen.add(number)
        normalized.append((number, to_excel_sheet_name(number)))
    if not normalized:
        raise DataValidationError("생성할 MKR 번호가 없습니다.")
    return normalized


def validate_target_sheets(workbook_path: Path, mkr_numbers: list[str]) -> dict[str, str]:
    """Verify that every preview target has a compatible existing sheet."""
    target_path = Path(workbook_path).resolve()
    normalized = _normalize_numbers(mkr_numbers)
    if not target_path.is_file():
        raise WorkbookCommitError(f"Excel 파일을 찾지 못했습니다: {target_path}")
    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    excel = workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(target_path), 0, True)
        names = _sheet_names(workbook)
        missing = [sheet for _, sheet in normalized if sheet not in names]
        if missing:
            raise DataValidationError("Outlook 분석 전에 MKR 시트를 생성해 주세요: " + ", ".join(missing))
        for number, sheet_name in normalized:
            _validate_existing_sheet(workbook.Worksheets.Item(sheet_name), number, None)
        return dict(normalized)
    except (DataValidationError, WorkbookCommitError):
        raise
    except Exception as exc:
        raise OfficeAutomationError(f"대상 Excel 시트 검증에 실패했습니다: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def create_mkr_sheets(
    workbook_path: Path,
    mkr_numbers: list[str],
    template_workbook: Path | None = None,
    template_name: str = TEMPLATE_SHEET,
    logger: logging.Logger | None = None,
    backup_dir: Path | None = None,
    progress: Callable[[int, str], None] | None = None,
) -> dict[str, object]:
    """Create missing MKR sheets transactionally from a separate read-only workbook."""
    log = logger or logging.getLogger("mkr_sync")
    target_path = Path(workbook_path).resolve()
    template_path = Path(template_workbook).resolve() if template_workbook else None
    if not target_path.is_file():
        raise WorkbookCommitError(f"Excel 파일을 찾지 못했습니다: {target_path}")
    if target_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise DataValidationError("대상 Excel은 .xlsx 또는 .xlsm 파일이어야 합니다.")
    normalized = _normalize_numbers(mkr_numbers)
    if template_path is not None and os.path.normcase(str(target_path)) == os.path.normcase(str(template_path)):
        raise DataValidationError("대상 Excel과 템플릿 Excel은 서로 다른 파일이어야 합니다.")

    assert_workbook_available(target_path)
    if progress:
        progress(5, "대상 Excel을 검사하고 있습니다.")

    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    excel = target = template_book = None
    existing: list[str] = []
    missing: list[tuple[str, str]] = []
    template_fingerprint: tuple | None = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        target = excel.Workbooks.Open(str(target_path), 0, True)
        names = _sheet_names(target)
        missing = [(number, sheet) for number, sheet in normalized if sheet not in names]

        if missing:
            if template_path is None or not template_path.is_file():
                raise DataValidationError("누락 시트를 만들 외부 MKR_TEMPLATE.xlsx가 필요합니다.")
            if template_path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
                raise DataValidationError("템플릿은 Excel 파일이어야 합니다.")
            template_book = excel.Workbooks.Open(str(template_path), 0, True)
            if template_name not in _sheet_names(template_book):
                raise DataValidationError(f"템플릿 파일에 {template_name} 시트가 없습니다.")
            template_sheet = template_book.Worksheets.Item(template_name)
            _validate_headers(template_sheet, template_name)
            _validate_no_external_formulas(template_sheet)
            template_fingerprint = _structure_fingerprint(template_sheet)

        for number, sheet_name in normalized:
            if sheet_name in names:
                _validate_existing_sheet(target.Worksheets.Item(sheet_name), number, template_fingerprint)
                existing.append(sheet_name)
    except (DataValidationError, WorkbookCommitError):
        raise
    except Exception as exc:
        raise OfficeAutomationError(f"Excel 사전 검사에 실패했습니다: {exc}") from exc
    finally:
        if template_book is not None:
            try:
                template_book.Close(False)
            except Exception:
                pass
        if target is not None:
            try:
                target.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not missing:
        return {"created": [], "existing": existing, "backup": None}

    backup = _backup_path(target_path, backup_dir)
    pending_path = target_path.with_name(
        f".{target_path.stem}_mkr_pending_{os.getpid()}{target_path.suffix}"
    )
    try:
        shutil.copy2(target_path, backup)
        shutil.copy2(target_path, pending_path)
    except OSError as exc:
        raise WorkbookCommitError(f"백업 또는 pending 파일을 만들지 못했습니다: {exc}") from exc

    if progress:
        progress(20, "백업을 완료하고 템플릿을 준비했습니다.")

    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    excel = target = template_book = None
    verified = False
    created_names: list[str] = []
    required_names = {sheet for _, sheet in normalized}
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        target = excel.Workbooks.Open(str(pending_path), 0, False)
        template_book = excel.Workbooks.Open(str(template_path), 0, True)
        template_sheet = template_book.Worksheets.Item(template_name)
        source_fingerprint = _structure_fingerprint(template_sheet)
        total = len(missing)
        for index, (number, sheet_name) in enumerate(missing, start=1):
            template_sheet.Copy(None, target.Worksheets.Item(int(target.Worksheets.Count)))
            created = target.Worksheets.Item(int(target.Worksheets.Count))
            created.Name = sheet_name
            created.Range("D3").Value2 = number
            if _structure_fingerprint(created) != source_fingerprint:
                raise WorkbookCommitError(f"복사된 시트 {sheet_name}의 구조 검증에 실패했습니다.")
            created_names.append(sheet_name)
            log.info("외부 템플릿에서 시트 생성: %s (%s)", sheet_name, number)
            if progress:
                progress(20 + int(index / total * 55), f"{sheet_name} 시트를 생성했습니다.")

        if template_name in _sheet_names(target):
            raise WorkbookCommitError("대상 Excel에 MKR_TEMPLATE 시트가 남았습니다.")
        target.Save()
        for number, sheet_name in normalized:
            if sheet_name not in _sheet_names(target):
                raise WorkbookCommitError(f"저장 후 시트를 찾지 못했습니다: {sheet_name}")
            _validate_existing_sheet(target.Worksheets.Item(sheet_name), number, source_fingerprint)
        verified = True
    except (DataValidationError, WorkbookCommitError):
        if pending_path.exists():
            try:
                pending_path.unlink()
            except OSError:
                pass
        raise
    except Exception as exc:
        if pending_path.exists():
            try:
                pending_path.unlink()
            except OSError:
                pass
        raise OfficeAutomationError(f"Excel 시트 생성에 실패했습니다: {exc}") from exc
    finally:
        if template_book is not None:
            try:
                template_book.Close(False)
            except Exception:
                pass
        if target is not None:
            try:
                target.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        if not verified and pending_path.exists():
            try:
                pending_path.unlink()
            except OSError:
                pass

    try:
        if not verified:
            raise WorkbookCommitError("pending 파일 검증이 완료되지 않았습니다.")
        disk_names = _disk_sheet_names(pending_path)
        if disk_names is not None:
            missing_on_disk = sorted(required_names.difference(disk_names))
            if missing_on_disk or template_name in disk_names:
                raise WorkbookCommitError(
                    "디스크 저장 후 시트 검증에 실패했습니다: "
                    + ", ".join(missing_on_disk or [template_name])
                )
        os.replace(pending_path, target_path)
        if progress:
            progress(100, "시트 생성과 원본 교체를 완료했습니다.")
        return {"created": created_names, "existing": existing, "backup": str(backup)}
    except Exception:
        if pending_path.exists():
            try:
                pending_path.unlink()
            except OSError:
                pass
        raise
