from __future__ import annotations

import logging
import threading
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .errors import MkrSyncError
from .jobs import JobCancelled
from .models import CollectedMail, MkrResult, PreviewRequest
from .mkr_parser import to_excel_sheet_name
from .outlook_service import collect_outlook_preview
from .office import sha256_file
from .reporting import write_audit_report
from .sheet_creator import validate_target_sheets
from .workflow import build_results


@dataclass(frozen=True)
class PreviewConfig:
    root: Path
    workbook: Path
    sender: str
    start_date: date
    max_messages: int
    data_start_row: int
    max_rows: int
    targets: dict[str, str]

    @property
    def data_last_row(self) -> int:
        return self.data_start_row + self.max_rows - 1


def _result_dict(result: MkrResult) -> dict:
    def quantity(value):
        return None if value is None else str(value)

    return {
        "mkr": result.mkr,
        "sheet": result.sheet,
        "order_date": result.order_date.isoformat() if result.order_date else None,
        "ship_date": result.ship_date.isoformat() if result.ship_date else None,
        "eta": result.eta.isoformat() if result.eta else None,
        "etd": result.etd.isoformat() if result.etd else None,
        "status": result.status,
        "warnings": list(result.warnings),
        "order_source": result.order_source,
        "sales_note_source": result.sales_note_source,
        "order_state_as_of": (
            result.order_state_as_of.isoformat() if result.order_state_as_of else None
        ),
        "new_order_count": len(result.decisions),
        "backorder_count": len(result.backorders),
        "decisions": [
            {
                "code": item.code,
                "name": item.name,
                "r": quantity(item.raw_order_qty),
                "s": quantity(item.first_shipped_qty),
                "p": quantity(item.sales_note_order_qty),
                "k": str(item.backorder_order_qty),
                "l": str(item.backorder_shipped_qty),
                "order_qty": quantity(item.order_qty),
                "shipped_qty": quantity(item.shipped_qty),
                "shortage_qty": quantity(item.shortage_qty),
                "rule": item.rule_id,
                "shipped_fallback": item.shipped_fallback,
                "shipped_adjusted": item.shipped_adjusted,
                "origin": item.origin,
                "record_status": item.record_status,
                "order_source": item.order_source,
                "sales_note_source": item.sales_note_source,
            }
            for item in result.decisions
        ],
        "backorders": [
            {
                "source_mkr": item.source_mkr,
                "code": item.code,
                "name": item.name,
                "order_qty": str(item.order_qty),
                "shipped_qty": str(item.shipped_qty),
                "shortage_qty": str(item.shortage_qty),
                "received_at": item.received_at.isoformat(),
                "source": item.source,
                "origin": item.origin,
            }
            for item in result.backorders
        ],
    }


def run_preview(
    root: Path,
    request: PreviewRequest,
    logger: logging.Logger,
    cancel_event: threading.Event,
    progress,
) -> dict:
    progress(3, "validate", "대상 Excel 시트를 검증하고 있습니다.")
    targets = validate_target_sheets(request.target_workbook, request.mkr_numbers)
    workbook_hash = sha256_file(request.target_workbook)
    if cancel_event.is_set():
        raise JobCancelled("사용자 요청으로 작업을 취소했습니다.")

    run_id = f"Run_{datetime.now():%Y%m%d_%H%M%S}_{threading.get_ident()}"
    attachment_dir = root / "MKR_Attachments" / run_id
    collected = collect_outlook_preview(
        request,
        attachment_dir,
        logger,
        cancel_event,
        progress=lambda value, message: progress(8 + int(value * 0.42), "outlook", message),
    )
    if cancel_event.is_set():
        raise JobCancelled("사용자 요청으로 작업을 취소했습니다.")

    results: list[MkrResult] = []
    errors: list[dict] = []
    warnings: list[dict] = []
    total = len(request.mkr_numbers)
    for index, mkr in enumerate(request.mkr_numbers, start=1):
        if cancel_event.is_set():
            raise JobCancelled("사용자 요청으로 작업을 취소했습니다.")
        progress(
            52 + int(index / max(total, 1) * 35),
            "analysis",
            f"{mkr} 주문서와 Sales Note를 분석하고 있습니다.",
        )
        subset = CollectedMail(
            records=[item for item in collected.records if item.target_mkr == mkr],
            attachments=[item for item in collected.attachments if item.target_mkr == mkr],
            scanned=collected.scanned,
        )
        config = PreviewConfig(
            root=root,
            workbook=request.target_workbook,
            sender=request.senders[0],
            start_date=request.start_date,
            max_messages=request.max_messages,
            data_start_row=10,
            max_rows=500,
            targets={mkr: targets[mkr]},
        )
        try:
            built = build_results(config, subset, logger)
            results.extend(built)
            for message in built[0].warnings if built else []:
                warnings.append(
                    {
                        "level": "warning",
                        "code": "REVIEW_RECOMMENDED",
                        "message": message,
                        "mkr": mkr,
                    }
                )
        except MkrSyncError as exc:
            logger.error("[%s] 분석 실패: %s", mkr, exc)
            errors.append(
                {
                    "level": "error",
                    "code": type(exc).__name__,
                    "message": str(exc),
                    "mkr": mkr,
                }
            )

    progress(92, "report", "CSV 감사 보고서를 생성하고 있습니다.")
    report = write_audit_report(root, results, dry_run=True, prefix="MKR_PREVIEW")
    return {
        "kind": "outlook_preview",
        "target_workbook": str(request.target_workbook.resolve()),
        "workbook_hash": workbook_hash,
        "mkr_numbers": list(request.mkr_numbers),
        "senders": list(request.senders),
        "start_date": request.start_date.isoformat(),
        "max_messages": request.max_messages,
        "results": [_result_dict(item) for item in results],
        "scanned": collected.scanned,
        "matched_messages": len(collected.records),
        "attachment_count": len(collected.attachments),
        "report_path": str(report),
        "attachment_dir": str(attachment_dir),
        "warnings": warnings,
        "errors": errors,
    }
