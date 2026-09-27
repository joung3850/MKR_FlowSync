from __future__ import annotations

import logging
import os
import threading
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .config import AppConfig
from .errors import DataValidationError
from .jobs import JobCancelled
from .models import BackorderSnapshot, MkrResult, QuantityDecision
from .mkr_parser import to_excel_sheet_name
from .office import WorkbookTransaction, sha256_file
from .reporting import write_audit_report


def _decimal(value: object, label: str) -> Decimal:
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise DataValidationError(f"미리보기 수량값이 올바르지 않습니다: {label}") from exc


def _date(value: object, label: str) -> date | None:
    if value in (None, ""):
        return None
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise DataValidationError(f"미리보기 날짜값이 올바르지 않습니다: {label}") from exc


def _datetime(value: object, label: str) -> datetime:
    try:
        return datetime.fromisoformat(str(value))
    except ValueError as exc:
        raise DataValidationError(f"미리보기 일시값이 올바르지 않습니다: {label}") from exc


def deserialize_preview_results(payload: dict) -> list[MkrResult]:
    raw_results = payload.get("results")
    if not isinstance(raw_results, list) or not raw_results:
        raise DataValidationError("Excel에 적용할 미리보기 결과가 없습니다.")

    results: list[MkrResult] = []
    seen_mkrs: set[str] = set()
    seen_sheets: set[str] = set()
    for raw_result in raw_results:
        if not isinstance(raw_result, dict):
            raise DataValidationError("미리보기 결과 형식이 올바르지 않습니다.")
        mkr = str(raw_result.get("mkr", "")).strip()
        sheet = str(raw_result.get("sheet", "")).strip()
        if not mkr or not sheet or sheet != to_excel_sheet_name(mkr):
            raise DataValidationError(f"미리보기 MKR과 시트명이 일치하지 않습니다: {mkr}/{sheet}")
        if mkr in seen_mkrs or sheet in seen_sheets:
            raise DataValidationError(f"미리보기 결과가 중복되었습니다: {mkr}")
        seen_mkrs.add(mkr)
        seen_sheets.add(sheet)

        decisions: list[QuantityDecision] = []
        for item in raw_result.get("decisions", []):
            if not isinstance(item, dict):
                raise DataValidationError(f"{mkr} 품목 결과 형식이 올바르지 않습니다.")
            code = str(item.get("code", "")).strip()
            name = str(item.get("name", "")).strip()
            if not code or not name:
                raise DataValidationError(f"{mkr} 품번 또는 품명이 비어 있습니다.")
            sales_note = item.get("p")
            decisions.append(
                QuantityDecision(
                    mkr=mkr,
                    code=code,
                    name=name,
                    raw_order_qty=_decimal(item.get("r"), f"{mkr}/{code}/R"),
                    first_shipped_qty=_decimal(item.get("s"), f"{mkr}/{code}/S"),
                    sales_note_order_qty=(
                        None
                        if sales_note in (None, "")
                        else _decimal(sales_note, f"{mkr}/{code}/P")
                    ),
                    backorder_order_qty=_decimal(item.get("k"), f"{mkr}/{code}/K"),
                    backorder_shipped_qty=_decimal(item.get("l"), f"{mkr}/{code}/L"),
                    order_qty=_decimal(item.get("order_qty"), f"{mkr}/{code}/Order"),
                    shipped_qty=_decimal(item.get("shipped_qty"), f"{mkr}/{code}/Shipped"),
                    rule_id=str(item.get("rule", "")).strip(),
                    shipped_fallback=bool(item.get("shipped_fallback", False)),
                    shipped_adjusted=bool(item.get("shipped_adjusted", False)),
                )
            )

        backorders: list[BackorderSnapshot] = []
        for item in raw_result.get("backorders", []):
            if not isinstance(item, dict):
                raise DataValidationError(f"{mkr} 백오더 결과 형식이 올바르지 않습니다.")
            source_mkr = str(item.get("source_mkr", "")).strip()
            code = str(item.get("code", "")).strip()
            name = str(item.get("name", "")).strip()
            if not source_mkr or not code or not name:
                raise DataValidationError(f"{mkr} 백오더 필수값이 비어 있습니다.")
            backorders.append(
                BackorderSnapshot(
                    current_mkr=mkr,
                    source_mkr=source_mkr,
                    code=code,
                    name=name,
                    order_qty=_decimal(item.get("order_qty"), f"{mkr}/{code}/Backorder Order"),
                    shipped_qty=_decimal(item.get("shipped_qty"), f"{mkr}/{code}/Backorder Shipped"),
                    received_at=_datetime(item.get("received_at"), f"{mkr}/{code}/Received"),
                    source=str(item.get("source", "")).strip(),
                )
            )

        results.append(
            MkrResult(
                mkr=mkr,
                sheet=sheet,
                decisions=decisions,
                backorders=backorders,
                order_date=_date(raw_result.get("order_date"), f"{mkr}/Order Date"),
                ship_date=_date(raw_result.get("ship_date"), f"{mkr}/Ship Date"),
                eta=_date(raw_result.get("eta"), f"{mkr}/ETA"),
                etd=_date(raw_result.get("etd"), f"{mkr}/ETD"),
            )
        )
    return results


def apply_preview_results(
    root: Path,
    target_workbook: Path,
    preview_payload: dict,
    logger: logging.Logger,
    cancel_event: threading.Event,
    progress,
) -> dict:
    if preview_payload.get("kind") != "outlook_preview":
        raise DataValidationError("Outlook 미리보기 결과가 아닙니다. 다시 분석해 주세요.")
    if preview_payload.get("errors"):
        raise DataValidationError("미리보기 오류가 남아 있어 Excel에 적용할 수 없습니다.")

    target = Path(target_workbook).resolve()
    preview_target = Path(str(preview_payload.get("target_workbook", ""))).resolve()
    if os.path.normcase(str(target)) != os.path.normcase(str(preview_target)):
        raise DataValidationError("미리보기 대상과 현재 선택한 Excel 파일이 다릅니다.")
    if not target.is_file():
        raise DataValidationError(f"대상 Excel 파일을 찾지 못했습니다: {target}")

    progress(5, "apply_validate", "미리보기 결과와 대상 Excel을 재검증하고 있습니다.")
    expected_hash = str(preview_payload.get("workbook_hash", "")).upper()
    if not expected_hash or sha256_file(target).upper() != expected_hash:
        raise DataValidationError(
            "미리보기 이후 대상 Excel이 변경되었습니다. 안전을 위해 다시 분석해 주세요."
        )
    results = deserialize_preview_results(preview_payload)
    if cancel_event.is_set():
        raise JobCancelled("사용자 요청으로 작업을 취소했습니다.")

    senders = preview_payload.get("senders") or ["preview@local.invalid"]
    config = AppConfig(
        root=Path(root).resolve(),
        workbook=target,
        sender=str(senders[0]),
        start_date=_date(preview_payload.get("start_date"), "검색 시작일") or date.today(),
        max_messages=int(preview_payload.get("max_messages", 500)),
        data_start_row=10,
        max_rows=500,
        targets={result.mkr: result.sheet for result in results},
    )
    transaction = WorkbookTransaction(config, logger)
    try:
        progress(15, "apply_backup", "대상 Excel의 실행 전 백업을 만들고 있습니다.")
        backup = transaction.prepare()
        if transaction.original_hash.upper() != expected_hash:
            raise DataValidationError(
                "백업 중 대상 Excel이 변경되었습니다. 다시 분석한 뒤 적용해 주세요."
            )
        progress(35, "apply_write", "분석 결과를 임시 Excel 파일에 기록하고 있습니다.")
        transaction.write_pending(results)
        progress(65, "apply_verify", "임시 Excel의 값과 구조를 검증하고 있습니다.")
        transaction.validate_pending(results)
        report = write_audit_report(
            root,
            results,
            dry_run=False,
            prefix="MKR_APPLIED",
        )
        if cancel_event.is_set():
            raise JobCancelled("원본 교체 전에 작업을 취소했습니다.")
        progress(88, "apply_commit", "검증된 Excel 파일로 안전하게 교체하고 있습니다.")
        transaction.commit()
        updated_hash = sha256_file(target)
        logger.info(
            "Excel 데이터 적용 완료: %s / 시트 %d개 / 신규 %d개 / 백오더 %d개",
            target,
            len(results),
            sum(len(item.decisions) for item in results),
            sum(len(item.backorders) for item in results),
        )
        return {
            "kind": "excel_apply",
            "target_workbook": str(target),
            "updated_sheets": [item.sheet for item in results],
            "new_order_count": sum(len(item.decisions) for item in results),
            "backorder_count": sum(len(item.backorders) for item in results),
            "backup_path": str(backup),
            "report_path": str(report),
            "workbook_hash": updated_hash,
            "warnings": [],
            "errors": [],
        }
    finally:
        transaction.cleanup_pending()
