from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from .config import AppConfig
from .errors import DataValidationError, MkrSyncError, OfficeAutomationError
from .models import BackorderSnapshot, CollectedMail, MkrResult
from .office import ExcelOrderReader, PdfTextReader
from .parsers import (
    merge_sales_evidence,
    parse_backorders,
    parse_date_near_keywords,
    parse_sales_note,
)
from .rules import build_decisions

def original_order_candidates(events):
    originals = sorted(
        (event for event in events if event.kind == "order" and not event.is_revision),
        key=lambda item: (item.received_at, item.original_name),
    )
    seen_hashes: set[str] = set()
    selected = []
    for event in originals:
        if event.sha256 in seen_hashes:
            continue
        seen_hashes.add(event.sha256)
        selected.append(event)
    return selected


def select_first_successful_order(parsed_events):
    """Match V11.8: earliest parseable order, then the most complete file."""
    if not parsed_events:
        return None
    first_received = min(event.received_at for event, _items in parsed_events)
    same_message_time = [
        (event, items)
        for event, items in parsed_events
        if event.received_at == first_received
    ]
    return sorted(
        same_message_time,
        key=lambda pair: (-len(pair[1]), pair[0].original_name.casefold()),
    )[0]


def merge_latest_backorders(
    ledger: dict[tuple[str, str], BackorderSnapshot],
    snapshot: dict[tuple[str, str], BackorderSnapshot],
    current_mkr: str,
) -> None:
    for key, item in snapshot.items():
        if item.source_mkr != current_mkr:
            ledger[key] = item


def validate_capacity(mkr: str, new_orders: int, backorders: int, maximum: int) -> None:
    if new_orders > maximum or backorders > maximum:
        raise DataValidationError(
            f"{mkr}의 출력 행이 {maximum}개를 초과했습니다: A:E={new_orders} / H:M={backorders}"
        )


def build_results(config: AppConfig, collected: CollectedMail, logger) -> list[MkrResult]:
    order_reader = ExcelOrderReader(logger)
    pdf_reader = PdfTextReader()
    pdf_cache: dict[Path, str] = {}
    results: list[MkrResult] = []

    records_by_target = defaultdict(list)
    attachments_by_target = defaultdict(list)
    for record in collected.records:
        records_by_target[record.target_mkr].append(record)
    for event in collected.attachments:
        attachments_by_target[event.target_mkr].append(event)

    def read_pdf(path: Path) -> str:
        if path not in pdf_cache:
            try:
                pdf_cache[path] = pdf_reader.read_text(path)
            except OfficeAutomationError as exc:
                logger.warning("%s", exc)
                pdf_cache[path] = ""
        return pdf_cache[path]

    for mkr, sheet_name in config.targets.items():
        records = sorted(records_by_target[mkr], key=lambda item: item.received_at)
        events = sorted(attachments_by_target[mkr], key=lambda item: (item.received_at, item.original_name))
        if not records and not events:
            logger.info("[%s] 대상 메일이 없어 기존 시트를 유지합니다.", mkr)
            continue

        candidates = original_order_candidates(events)
        if not candidates:
            raise DataValidationError(
                f"{mkr}의 최초 원주문 첨부를 찾지 못했습니다. "
                f"검색 시작일({config.start_date.isoformat()})을 원주문 메일 날짜 이전으로 설정해 주세요."
            )

        parsed_originals = []
        for event in candidates:
            try:
                parsed = order_reader.read(event.path, mkr)
            except MkrSyncError as exc:
                logger.warning(
                    "[%s] 원본 주문 후보 건너뜀: %s / %s",
                    mkr,
                    event.original_name,
                    exc,
                )
                continue
            logger.info("[%s] 원본 주문 후보 분석: %s / %d개", mkr, event.original_name, len(parsed))
            if parsed:
                parsed_originals.append((event, parsed))

        selected_order = select_first_successful_order(parsed_originals)
        if selected_order is None:
            raise DataValidationError(
                f"{mkr}의 원주문 후보에서 품목을 읽지 못했습니다. "
                "검색 시작일과 원주문 첨부파일을 확인해 주세요."
            )
        first_event, baseline = selected_order
        logger.info(
            "[%s] 최초 유효 원본 주문 확정: %s / %d개",
            mkr,
            first_event.original_name,
            len(baseline),
        )
        validate_capacity(mkr, len(baseline), 0, config.max_rows)

        mail_by_entry = {record.entry_id: record for record in records}
        normal_sales = [event for event in events if event.kind == "sales_note" and not event.is_revision]
        sales_evidence = {}
        ship_date = eta = etd = None
        if normal_sales:
            first_sales = normal_sales[0]
            first_sales_evidence = {}
            for candidate in normal_sales:
                pdf_evidence = parse_sales_note(read_pdf(candidate.path))
                body_evidence = parse_sales_note(candidate.body)
                candidate_evidence = merge_sales_evidence(pdf_evidence, body_evidence)
                logger.info(
                    "[%s] 일반 SALES NOTE 후보 분석: %s / %d개 품목",
                    mkr,
                    candidate.original_name,
                    len(candidate_evidence),
                )
                if candidate_evidence:
                    first_sales = candidate
                    first_sales_evidence = candidate_evidence
                    break
            sales_evidence = first_sales_evidence
            source_record = mail_by_entry.get(first_sales.entry_id)
            if source_record:
                reference = source_record.received_at
                ship_date = parse_date_near_keywords(
                    source_record.body,
                    r"\b(?:Ship|Shipment|Shipping)\s*Date\b|出荷日|発送日|출하일|선적일",
                    reference,
                )
                eta = parse_date_near_keywords(source_record.body, r"\bETA\b|到着予定|着港|도착예정", reference)
                etd = parse_date_near_keywords(source_record.body, r"\bETD\b|出港|船積|출항", reference)
            logger.info(
                "[%s] 최초 일반 SALES NOTE 확정: %s / %d개 품목",
                mkr,
                first_sales.original_name,
                len(sales_evidence),
            )
        else:
            logger.warning("[%s] 일반 SALES NOTE가 없어 Shipped QTY fallback을 적용합니다.", mkr)

        # Follow-up schedule messages override earlier values only when they
        # contain an explicit date. This keeps the first Sales Note as the
        # quantity baseline while allowing later ETA/ETD updates.
        for record in records:
            reference = record.received_at
            ship_date = parse_date_near_keywords(
                record.body,
                r"\b(?:Ship|Shipment|Shipping)\s*Date\b|出荷日|発送日|출하일|선적일",
                reference,
            ) or ship_date
            eta = parse_date_near_keywords(
                record.body, r"\bETA\b|到着予定|着港|도착예정", reference
            ) or eta
            etd = parse_date_near_keywords(
                record.body, r"\bETD\b|出港|船積|출항", reference
            ) or etd

        sales_by_entry = defaultdict(list)
        for event in events:
            if event.kind == "sales_note":
                sales_by_entry[event.entry_id].append(event)

        ledger: dict[tuple[str, str], BackorderSnapshot] = {}
        for record in records:
            event_snapshot = parse_backorders(
                record.body,
                mkr,
                record.received_at,
                f"메일 본문 {record.received_at:%Y-%m-%d %H:%M}",
            )
            seen_pdf_hashes: set[str] = set()
            for event in sales_by_entry[record.entry_id]:
                if event.sha256 in seen_pdf_hashes:
                    continue
                seen_pdf_hashes.add(event.sha256)
                pdf_snapshot = parse_backorders(
                    read_pdf(event.path),
                    mkr,
                    event.received_at,
                    f"SALES NOTE {event.original_name}",
                )
                event_snapshot.update(pdf_snapshot)
            merge_latest_backorders(ledger, event_snapshot, mkr)

        backorders = sorted(ledger.values(), key=lambda item: (item.source_mkr, item.code))
        validate_capacity(mkr, len(baseline), len(backorders), config.max_rows)
        decisions = build_decisions(mkr, baseline, sales_evidence, backorders)
        results.append(
            MkrResult(
                mkr=mkr,
                sheet=sheet_name,
                decisions=decisions,
                backorders=backorders,
                order_date=first_event.received_at.date(),
                ship_date=ship_date,
                eta=eta,
                etd=etd,
            )
        )
        logger.info("[%s] 판정 완료: 신규 %d개 / 백오더 %d개", mkr, len(decisions), len(backorders))

    if not results:
        raise DataValidationError("읽을 수 있는 최초 주문 자료가 없어 갱신할 시트가 없습니다.")
    return results
