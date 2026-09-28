from __future__ import annotations

import ctypes
import hashlib
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from .config import AppConfig
from .errors import DataValidationError, OfficeAutomationError, WorkbookCommitError
from .models import AttachmentEvent, CollectedMail, MailRecord, MkrResult, OrderItem
from .parsers import (
    attachment_kind,
    clean_text,
    current_message_body,
    find_all_mkrs,
    find_target_mkrs,
    parse_order_matrix,
    validate_mkr_match,
)


def _load_com():
    try:
        import pythoncom  # type: ignore
        # Outlook COM date values can be unpickled through this module at
        # runtime, so keep the import explicit for frozen PyInstaller builds.
        import win32timezone  # type: ignore  # noqa: F401
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise OfficeAutomationError(
            "pywin32가 설치되지 않았습니다. SETUP_MKR.cmd를 먼저 실행하세요."
        ) from exc
    return pythoncom, win32com.client


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_filename(value: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", value).strip()
    return name[:180] or "attachment"


def _sender_smtp(mail) -> str:
    try:
        if clean_text(mail.SenderEmailType).upper() == "EX":
            sender = mail.Sender
            if sender is not None:
                exchange_user = sender.GetExchangeUser()
                if exchange_user is not None and clean_text(exchange_user.PrimarySmtpAddress):
                    return clean_text(exchange_user.PrimarySmtpAddress).lower()
    except Exception:
        pass
    try:
        return clean_text(mail.SenderEmailAddress).lower()
    except Exception:
        return ""


def collect_outlook(config: AppConfig, run_attachment_dir: Path, logger) -> CollectedMail:
    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    result = CollectedMail()
    run_attachment_dir.mkdir(parents=True, exist_ok=True)
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception as exc:
            raise OfficeAutomationError(
                "실행 중인 Classic Outlook을 찾지 못했습니다. Classic Outlook을 먼저 여세요."
            ) from exc
        namespace = outlook.Session
        inbox = namespace.GetDefaultFolder(6)
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        count = min(int(items.Count), config.max_messages)
        for index in range(1, count + 1):
            mail = items.Item(index)
            result.scanned += 1
            try:
                if int(mail.Class) != 43:
                    continue
                received = mail.ReceivedTime
                if not isinstance(received, datetime):
                    received = datetime.fromisoformat(str(received))
                if received.date() < config.start_date:
                    break
                if _sender_smtp(mail) != config.sender:
                    continue
                subject = clean_text(mail.Subject)
                target_keys = find_target_mkrs(subject, config.targets)
                if not target_keys:
                    continue
                if len(target_keys) != 1:
                    raise DataValidationError(
                        f"메일 제목에서 여러 대상 MKR이 발견되었습니다: {subject} / {', '.join(target_keys)}"
                    )
                target = target_keys[0]
                body = current_message_body(mail.Body)
                entry_id = clean_text(mail.EntryID) or f"mail-{index}-{received:%Y%m%d%H%M%S}"
                result.records.append(MailRecord(entry_id, target, received, subject, body))

                for attachment_index in range(1, int(mail.Attachments.Count) + 1):
                    attachment = mail.Attachments.Item(attachment_index)
                    filename = clean_text(attachment.FileName)
                    kind, revision = attachment_kind(filename, subject)
                    if kind is None:
                        continue
                    file_mkrs = find_all_mkrs(filename)
                    validate_mkr_match(target, file_mkrs, f"첨부파일 {filename}")
                    destination = run_attachment_dir / (
                        f"{received:%Y%m%d_%H%M%S}_{index:03d}_{attachment_index:02d}_{_safe_filename(filename)}"
                    )
                    attachment.SaveAsFile(str(destination))
                    result.attachments.append(
                        AttachmentEvent(
                            entry_id=entry_id,
                            target_mkr=target,
                            received_at=received,
                            subject=subject,
                            body=body,
                            kind=kind,
                            is_revision=revision,
                            original_name=filename,
                            path=destination,
                            sha256=sha256_file(destination),
                        )
                    )
                    logger.info("첨부 저장: %s / %s / %s", target, kind, filename)
                logger.info("대상 메일: %s / %s / %s", received.isoformat(sep=" "), target, subject)
            finally:
                del mail
        if not result.records:
            raise DataValidationError("조건에 맞는 대상 MKR 메일을 찾지 못했습니다.")
        return result
    finally:
        pythoncom.CoUninitialize()


class WordPdfReader:
    def read_text(self, path: Path) -> str:
        pythoncom, win32 = _load_com()
        pythoncom.CoInitialize()
        word = document = None
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(str(path), False, True, False)
            return clean_text(document.Content.Text)
        except Exception as exc:
            raise OfficeAutomationError(f"SALES NOTE PDF를 읽지 못했습니다: {path.name} / {exc}") from exc
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


class PdfTextReader:
    """Extract embedded PDF text first and use Word only as a compatibility fallback."""

    def read_text(self, path: Path) -> str:
        try:
            from pypdf import PdfReader  # type: ignore

            reader = PdfReader(str(path))
            text = "\n".join((page.extract_text() or "") for page in reader.pages)
            if clean_text(text):
                return clean_text(text)
        except Exception:
            pass
        return WordPdfReader().read_text(path)


class ExcelOrderReader:
    def __init__(self, logger):
        self.logger = logger

    def read(self, path: Path, expected_mkr: str) -> dict[str, OrderItem]:
        pythoncom, win32 = _load_com()
        pythoncom.CoInitialize()
        excel = workbook = None
        combined: dict[str, OrderItem] = {}
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(path), 0, True)
            discovered: set[str] = set()
            for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
                sheet = workbook.Worksheets.Item(sheet_index)
                used = sheet.UsedRange
                matrix = _as_matrix(used.Value2)
                for row in matrix[:80]:
                    for value in row[:30]:
                        discovered.update(find_all_mkrs(value))
                parsed, used_fallback = parse_order_matrix(matrix)
                if used_fallback and parsed:
                    self.logger.warning("주문서 헤더 fallback 사용: %s / %s", path.name, sheet.Name)
                for code, item in parsed.items():
                    previous = combined.get(code)
                    combined[code] = (
                        OrderItem(
                            code,
                            previous.name or item.name,
                            previous.quantity + item.quantity,
                            previous.backorders_separated or item.backorders_separated,
                        )
                        if previous
                        else item
                    )
            validate_mkr_match(expected_mkr, discovered, f"주문서 {path.name}")
            return combined
        except DataValidationError:
            raise
        except Exception as exc:
            raise OfficeAutomationError(f"주문 Excel을 읽지 못했습니다: {path.name} / {exc}") from exc
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


def _as_matrix(value) -> list[list[object]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _decimal_for_excel(value: Decimal) -> int | float:
    if value == value.to_integral_value():
        return int(value)
    return float(value)


def _matrix(value) -> tuple[tuple[object, ...], ...]:
    return tuple(tuple(row) for row in value)


def _is_workbook_open(path: Path) -> bool:
    try:
        _, win32 = _load_com()
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        return False
    target = os.path.normcase(os.path.abspath(path))
    try:
        for index in range(1, int(excel.Workbooks.Count) + 1):
            candidate = os.path.normcase(os.path.abspath(clean_text(excel.Workbooks.Item(index).FullName)))
            if candidate == target:
                return True
    except Exception:
        return True
    return False


def _has_exclusive_file_access(path: Path) -> bool:
    if os.name != "nt":
        return True
    kernel32 = ctypes.windll.kernel32
    from ctypes import wintypes

    kernel32.CreateFileW.argtypes = (
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    )
    kernel32.CreateFileW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL
    handle = kernel32.CreateFileW(
        str(path),
        0x80000000 | 0x40000000,
        0,
        None,
        3,
        0x80,
        None,
    )
    if handle == wintypes.HANDLE(-1).value:
        return False
    kernel32.CloseHandle(handle)
    return True


def assert_workbook_available(path: Path) -> None:
    if _is_workbook_open(path) or not _has_exclusive_file_access(path):
        raise WorkbookCommitError(f"Excel 파일이 열려 있거나 잠겨 있습니다: {path.name}")


def _xlsx_merge_map(path: Path) -> dict[str, tuple[str, ...]]:
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    result: dict[str, tuple[str, ...]] = {}
    with zipfile.ZipFile(path) as archive:
        for name in sorted(item for item in archive.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", item)):
            root = ET.fromstring(archive.read(name))
            refs = tuple(
                sorted(node.attrib["ref"] for node in root.findall(".//x:mergeCells/x:mergeCell", namespace))
            )
            result[name] = refs
    return result


def _validate_preserved_number_formats(
    sheet_name: str,
    source_formats: dict[str, str],
    pending_formats: dict[str, str],
) -> None:
    for address, expected_format in source_formats.items():
        actual_format = pending_formats.get(address, "")
        if actual_format != expected_format:
            raise WorkbookCommitError(
                f"{sheet_name}의 셀 서식이 변경되었습니다: {address} "
                f"(원본={expected_format or '없음'}, 임시={actual_format or '없음'})"
            )


class WorkbookTransaction:
    def __init__(self, config: AppConfig, logger):
        self.config = config
        self.logger = logger
        self.original_hash = ""
        self.backup_path: Path | None = None
        self.source_number_formats: dict[str, dict[str, str]] = {}
        self.pending_path = config.root / (
            f".{config.workbook.stem}_data_pending{config.workbook.suffix}"
        )

    def prepare(self) -> Path:
        try:
            assert_workbook_available(self.config.workbook)
            self.original_hash = sha256_file(self.config.workbook)
            backup_dir = self.config.root / "Backup"
            backup_dir.mkdir(parents=True, exist_ok=True)
            self.backup_path = backup_dir / (
                f"{self.config.workbook.stem}_before_data_apply_"
                f"{datetime.now():%Y%m%d_%H%M%S_%f}{self.config.workbook.suffix}"
            )
            shutil.copy2(self.config.workbook, self.backup_path)
            if self.pending_path.exists():
                self.pending_path.unlink()
            self.logger.info("실행 전 백업 생성: %s", self.backup_path)
            return self.backup_path
        except WorkbookCommitError:
            raise
        except OSError as exc:
            raise WorkbookCommitError(f"실행 전 백업을 만들지 못했습니다: {exc}") from exc

    def write_pending(self, results: Iterable[MkrResult]) -> None:
        pythoncom, win32 = _load_com()
        pythoncom.CoInitialize()
        excel = workbook = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(self.config.workbook), 0, False)
            for result in results:
                sheet = workbook.Worksheets.Item(result.sheet)
                self.source_number_formats[result.sheet] = self._data_number_formats(sheet)
                self._write_sheet(sheet, result)
            workbook.SaveCopyAs(str(self.pending_path))
        except Exception as exc:
            raise WorkbookCommitError(f"임시 Excel 저장에 실패했습니다: {exc}") from exc
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

    def _write_sheet(self, sheet, result: MkrResult) -> None:
        if len(result.decisions) > self.config.max_rows or len(result.backorders) > self.config.max_rows:
            raise DataValidationError(
                f"{result.mkr}의 출력 행이 500개를 초과했습니다: "
                f"A:E={len(result.decisions)} / H:M={len(result.backorders)}"
            )
        first, last = self.config.data_start_row, self.config.data_last_row
        sheet.Range(f"A{first}:E{last}").ClearContents()
        sheet.Range(f"H{first}:M{last}").ClearContents()

        if result.decisions:
            end = first + len(result.decisions) - 1
            text_values = _matrix((d.code, d.name) for d in result.decisions)
            number_values = _matrix(
                (
                    _decimal_for_excel(d.order_qty),
                    _decimal_for_excel(d.shipped_qty),
                    _decimal_for_excel(d.shortage_qty),
                )
                for d in result.decisions
            )
            sheet.Range(f"A{first}:B{end}").NumberFormat = "@"
            sheet.Range(f"A{first}:B{end}").Value2 = text_values
            sheet.Range(f"C{first}:E{end}").Value2 = number_values

        if result.backorders:
            end = first + len(result.backorders) - 1
            text_values = _matrix((b.source_mkr, b.code, b.name) for b in result.backorders)
            number_values = _matrix(
                (
                    _decimal_for_excel(b.order_qty),
                    _decimal_for_excel(b.shipped_qty),
                    _decimal_for_excel(b.shortage_qty),
                )
                for b in result.backorders
            )
            sheet.Range(f"H{first}:J{end}").NumberFormat = "@"
            sheet.Range(f"H{first}:J{end}").Value2 = text_values
            sheet.Range(f"K{first}:M{end}").Value2 = number_values

        sheet.Range("D3").Value2 = result.mkr
        sheet.Range("K4:K7").ClearContents()
        for address, value in (("K4", result.order_date), ("K5", result.ship_date), ("K6", result.eta), ("K7", result.etd)):
            if value is not None:
                sheet.Range(address).NumberFormat = "@"
                sheet.Range(address).Value2 = value.isoformat()
        self.logger.info(
            "시트 임시 반영: %s / A:E %d개 / H:M %d개",
            result.sheet,
            len(result.decisions),
            len(result.backorders),
        )

    def validate_pending(self, results: Iterable[MkrResult]) -> None:
        if not self.pending_path.is_file():
            raise WorkbookCommitError("임시 Excel 파일이 생성되지 않았습니다.")
        required_entries = {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            *(f"xl/worksheets/sheet{index}.xml" for index in range(1, 6)),
        }
        with zipfile.ZipFile(self.pending_path) as archive:
            missing = required_entries.difference(archive.namelist())
            if missing:
                raise WorkbookCommitError(f"XLSX 필수 항목이 누락되었습니다: {', '.join(sorted(missing))}")
        if _xlsx_merge_map(self.config.workbook) != _xlsx_merge_map(self.pending_path):
            raise WorkbookCommitError("통합문서 병합 셀 구조가 변경되었습니다.")

        pythoncom, win32 = _load_com()
        pythoncom.CoInitialize()
        excel = workbook = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(self.pending_path), 0, True)
            sheet_names = {clean_text(workbook.Worksheets.Item(i).Name) for i in range(1, int(workbook.Worksheets.Count) + 1)}
            missing_sheets = set(self.config.targets.values()).difference(sheet_names)
            if missing_sheets:
                raise WorkbookCommitError(f"필수 시트가 누락되었습니다: {', '.join(sorted(missing_sheets))}")
            for sheet_name in self.config.targets.values():
                formula_cells = workbook.Worksheets.Item(sheet_name).Range(
                    f"A1:R{self.config.data_last_row}"
                ).Formula
                for row in _as_matrix(formula_cells):
                    if any(isinstance(value, str) and value.startswith("=") for value in row):
                        raise WorkbookCommitError(f"{sheet_name}에 예상하지 않은 Excel 수식이 있습니다.")
            for result in results:
                self._validate_sheet(workbook.Worksheets.Item(result.sheet), result)
        except WorkbookCommitError:
            raise
        except Exception as exc:
            raise WorkbookCommitError(f"임시 Excel 검증에 실패했습니다: {exc}") from exc
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

    def _validate_sheet(self, sheet, result: MkrResult) -> None:
        headers = tuple(sheet.Range("A9:M9").Value2[0])
        expected = {
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
        for index, label in expected.items():
            if clean_text(headers[index]) != label:
                raise WorkbookCommitError(f"{result.sheet} 헤더가 다릅니다: {index + 1}열")
        if clean_text(sheet.Range("D3").Value2) != result.mkr:
            raise WorkbookCommitError(f"{result.sheet}의 MKR 번호가 다릅니다.")
        source_formats = self.source_number_formats.get(result.sheet)
        if source_formats is None:
            raise WorkbookCommitError(f"{result.sheet}의 원본 셀 서식 기준값이 없습니다.")
        pending_formats = self._data_number_formats(sheet)
        _validate_preserved_number_formats(result.sheet, source_formats, pending_formats)
        first = self.config.data_start_row
        for offset, decision in enumerate(result.decisions):
            row = first + offset
            values = tuple(sheet.Range(f"A{row}:E{row}").Value2[0])
            if clean_text(values[0]).zfill(6) != decision.code:
                raise WorkbookCommitError(f"{result.sheet} A:E 품번 검증 실패: {decision.code}")
            actual = tuple(Decimal(str(values[i])) for i in range(2, 5))
            expected_qty = (decision.order_qty, decision.shipped_qty, decision.shortage_qty)
            if actual != expected_qty:
                raise WorkbookCommitError(f"{result.sheet} A:E 수량 검증 실패: {decision.code}")
        for offset, backorder in enumerate(result.backorders):
            row = first + offset
            values = tuple(sheet.Range(f"H{row}:M{row}").Value2[0])
            if clean_text(values[0]) != backorder.source_mkr or clean_text(values[1]).zfill(6) != backorder.code:
                raise WorkbookCommitError(f"{result.sheet} H:M 키 검증 실패: {backorder.source_mkr}/{backorder.code}")
            actual = tuple(Decimal(str(values[i])) for i in range(3, 6))
            expected_qty = (backorder.order_qty, backorder.shipped_qty, backorder.shortage_qty)
            if actual != expected_qty:
                raise WorkbookCommitError(f"{result.sheet} H:M 수량 검증 실패: {backorder.code}")

    def _data_number_formats(self, sheet) -> dict[str, str]:
        first = self.config.data_start_row
        addresses = (
            f"A{first}",
            f"C{first}",
            f"D{first}",
            f"E{first}",
            f"H{first}",
            f"K{first}",
            f"L{first}",
            f"M{first}",
        )
        return {
            address: clean_text(sheet.Range(address).NumberFormat)
            for address in addresses
        }
    def commit(self) -> None:
        try:
            if sha256_file(self.config.workbook) != self.original_hash:
                raise WorkbookCommitError("실행 중 원본 Excel이 변경되어 교체를 중단했습니다.")
            assert_workbook_available(self.config.workbook)
            os.replace(self.pending_path, self.config.workbook)
            self.logger.info("검증된 Excel로 원본 교체 완료: %s", self.config.workbook)
        except WorkbookCommitError:
            raise
        except OSError as exc:
            raise WorkbookCommitError(f"검증된 Excel로 교체하지 못했습니다: {exc}") from exc

    def cleanup_pending(self) -> None:
        if self.pending_path.exists():
            try:
                self.pending_path.unlink()
            except OSError:
                pass
