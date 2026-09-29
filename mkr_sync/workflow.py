from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, replace
from pathlib import Path

from .config import AppConfig
from .errors import DataValidationError, MkrSyncError, OfficeAutomationError
from .models import AttachmentEvent, BackorderSnapshot, CollectedMail, MkrResult, OrderItem
from .office import ExcelOrderReader, PdfTextReader
from .parsers import merge_sales_evidence, parse_backorders, parse_date_near_keywords, parse_sales_note
from .rules import build_decisions


@dataclass(frozen=True)
class OrderState:
    items: dict[str, OrderItem]
    first_event: AttachmentEvent
    latest_event: AttachmentEvent
    warnings: tuple[str, ...] = ()


def original_order_candidates(events):
    """Return every unique order version, including revisions and additions."""
    ordered = sorted(
        (event for event in events if event.kind == "order"),
        key=lambda item: (item.received_at, item.entry_id, item.original_name.casefold()),
    )
    seen_hashes: set[str] = set()
    selected = []
    for event in ordered:
        if event.sha256 in seen_hashes:
            continue
        seen_hashes.add(event.sha256)
        selected.append(event)
    return selected


def select_first_successful_order(parsed_events):
    """Legacy helper retained for callers; reconstruction no longer stops here."""
    if not parsed_events:
        return None
    first_received = min(event.received_at for event, _items in parsed_events)
    same_message_time = [
        (event, items) for event, items in parsed_events if event.received_at == first_received
    ]
    return sorted(
        same_message_time,
        key=lambda pair: (-len(pair[1]), pair[0].original_name.casefold()),
    )[0]


def _origin_from_name(name: str) -> str:
    normalized = name.upper()
    if "THB" in normalized or re.search(r"(?:^|[\s_\-([])B(?:[\s_\-.)\]]|$)", normalized):
        return "Thailand"
    if "KRW" in normalized or re.search(r"(?:^|[\s_\-([])A(?:[\s_\-.)\]]|$)", normalized):
        return "Japan"
    return "Unknown"


def _combine_origin(left: str, right: str) -> str:
    known = {value for value in (left, right) if value and value != "Unknown"}
    if len(known) > 1:
        return "Mixed"
    return next(iter(known), "Unknown")


def _combine_source(left: str, right: str) -> str:
    parts = []
    for value in (left, right):
        for part in filter(None, (item.strip() for item in value.split(" + "))):
            if part not in parts:
                parts.append(part)
    return " + ".join(parts)


def merge_order_group(
    parsed_events: list[tuple[AttachmentEvent, dict[str, OrderItem]]],
) -> dict[str, OrderItem]:
    """Merge country/part files from one mail into one unambiguous snapshot."""
    combined: dict[str, OrderItem] = {}
    for event, parsed in parsed_events:
        origin = _origin_from_name(event.original_name)
        for code, raw_item in parsed.items():
            item = replace(raw_item, origin=origin, source=event.original_name)
            previous = combined.get(code)
            if previous is None:
                combined[code] = item
                continue
            if (
                previous.origin != item.origin
                and previous.origin != "Unknown"
                and item.origin != "Unknown"
            ):
                combined[code] = replace(
                    previous,
                    name=previous.name or item.name,
                    quantity=previous.quantity + item.quantity,
                    backorders_separated=(
                        previous.backorders_separated or item.backorders_separated
                    ),
                    origin="Mixed",
                    source=_combine_source(previous.source, item.source),
                )
                continue
            if previous.quantity != item.quantity:
                raise DataValidationError(
                    f"같은 메일의 주문 첨부에서 {code} 수량이 충돌합니다: "
                    f"{previous.quantity} / {item.quantity}"
                )
            combined[code] = replace(
                previous,
                name=previous.name or item.name,
                backorders_separated=previous.backorders_separated or item.backorders_separated,
                origin=_combine_origin(previous.origin, item.origin),
                source=_combine_source(previous.source, item.source),
            )
    return combined


def _event_role(events: list[AttachmentEvent]) -> str:
    text = " ".join(f"{event.subject} {event.original_name}" for event in events)
    if re.search(r"追加|\badditional\b|\badd(?:ition)?\b", text, re.IGNORECASE):
        return "addition"
    if any(event.is_revision for event in events) or re.search(
        r"\brev(?:ision)?\s*\d*\b|改訂|修正|수정|변경", text, re.IGNORECASE
    ):
        return "revision"
    return "ordinary"


def _same_quantities(left: dict[str, OrderItem], right: dict[str, OrderItem]) -> bool:
    return all(code in left and left[code].quantity == item.quantity for code, item in right.items())


def apply_order_version(
    current: dict[str, OrderItem], incoming: dict[str, OrderItem], role: str
) -> tuple[dict[str, OrderItem], str | None]:
    """Apply one chronological order event without inventing ambiguous cancellations."""
    if not current:
        return dict(incoming), None
    current_codes, incoming_codes = set(current), set(incoming)
    if role == "revision":
        return dict(incoming), None
    if role == "addition":
        if current_codes <= incoming_codes:
            return dict(incoming), None
        if current_codes.isdisjoint(incoming_codes):
            return {**current, **incoming}, None
        raise DataValidationError(
            "추가 주문 파일이 기존 일부 품목만 포함하여 누적본인지 증분본인지 판단할 수 없습니다."
        )
    if current_codes <= incoming_codes:
        return dict(incoming), None
    if current_codes.isdisjoint(incoming_codes):
        return {**current, **incoming}, None
    if incoming_codes < current_codes and _same_quantities(current, incoming):
        return dict(current), "기존 주문의 일부만 반복된 첨부는 물류용 사본으로 보고 제외했습니다."
    if current_codes == incoming_codes and _same_quantities(current, incoming):
        return dict(incoming), None
    raise DataValidationError(
        "일반 주문 파일이 기존 상태와 부분적으로 충돌하여 자동 병합하지 않았습니다."
    )


def reconstruct_order_state(events, order_reader, mkr: str, logger) -> OrderState:
    candidates = original_order_candidates(events)
    if not candidates:
        raise DataValidationError(f"{mkr}의 주문 첨부를 찾지 못했습니다.")
    grouped: dict[str, list[AttachmentEvent]] = defaultdict(list)
    for event in candidates:
        key = event.entry_id or f"{event.received_at.isoformat()}::{event.original_name}"
        grouped[key].append(event)
    ordered_groups = sorted(
        grouped.values(),
        key=lambda group: (min(item.received_at for item in group), min(item.original_name for item in group)),
    )
    state: dict[str, OrderItem] = {}
    first_event = latest_event = None
    warnings: list[str] = []
    for group in ordered_groups:
        parsed_group: list[tuple[AttachmentEvent, dict[str, OrderItem]]] = []
        for event in sorted(group, key=lambda item: item.original_name.casefold()):
            try:
                parsed = order_reader.read(event.path, mkr)
            except MkrSyncError as exc:
                message = f"주문 첨부 분석 제외: {event.original_name} ({exc})"
                logger.warning("[%s] %s", mkr, message)
                warnings.append(message)
                continue
            if parsed:
                parsed_group.append((event, parsed))
                logger.info("[%s] 주문 버전 분석: %s / %d개", mkr, event.original_name, len(parsed))
        if not parsed_group:
            continue
        incoming = merge_order_group(parsed_group)
        role = _event_role([event for event, _items in parsed_group])
        state, warning = apply_order_version(state, incoming, role)
        applied_event = max((event for event, _items in parsed_group), key=lambda item: item.received_at)
        first_event = first_event or min(
            (event for event, _items in parsed_group), key=lambda item: item.received_at
        )
        latest_event = applied_event
        if warning:
            warnings.append(f"{applied_event.original_name}: {warning}")
        logger.info("[%s] 주문 상태 반영: %s / %s / %d개", mkr, applied_event.original_name, role, len(state))
    if not state or first_event is None or latest_event is None:
        raise DataValidationError(f"{mkr} 주문 첨부에서 품목을 읽지 못했습니다.")
    return OrderState(state, first_event, latest_event, tuple(warnings))


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
            f"{mkr}의 출력 행이 {maximum}개를 초과했습니다: A:G={new_orders} / H:N={backorders}"
        )


def _latest_sales_state(events, records, read_pdf, mkr: str, logger):
    sales_events = sorted(
        (event for event in events if event.kind == "sales_note"),
        key=lambda item: (item.received_at, item.entry_id, item.original_name.casefold()),
    )
    record_by_entry = {record.entry_id: record for record in records}
    grouped: dict[str, list[AttachmentEvent]] = defaultdict(list)
    for event in sales_events:
        key = event.entry_id or f"{event.received_at.isoformat()}::{event.original_name}"
        grouped[key].append(event)
    selected = {}
    selected_event = None
    warnings: list[str] = []
    for group in sorted(grouped.values(), key=lambda value: min(item.received_at for item in value)):
        group_evidence = {}
        seen_hashes: set[str] = set()
        for event in group:
            if event.sha256 in seen_hashes:
                continue
            seen_hashes.add(event.sha256)
            group_evidence = merge_sales_evidence(group_evidence, parse_sales_note(read_pdf(event.path)))
        record = record_by_entry.get(group[0].entry_id)
        body = record.body if record is not None else group[0].body
        group_evidence = merge_sales_evidence(group_evidence, parse_sales_note(body))
        latest = max(group, key=lambda item: item.received_at)
        if group_evidence:
            selected = group_evidence
            selected_event = latest
            logger.info("[%s] SALES NOTE 상태 반영: %s / %d개", mkr, latest.original_name, len(selected))
        else:
            warnings.append(f"SALES NOTE 분석 제외: {latest.original_name} (품목을 찾지 못함)")
    return selected, selected_event, warnings


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
        events = sorted(
            attachments_by_target[mkr],
            key=lambda item: (item.received_at, item.entry_id, item.original_name.casefold()),
        )
        if not records and not events:
            logger.info("[%s] 대상 메일이 없어 기존 시트를 유지합니다.", mkr)
            continue
        order_state = reconstruct_order_state(events, order_reader, mkr, logger)
        warnings = list(order_state.warnings)
        sales_evidence, sales_event, sales_warnings = _latest_sales_state(
            events, records, read_pdf, mkr, logger
        )
        warnings.extend(sales_warnings)

        ship_date = eta = etd = None
        for record in records:
            reference = record.received_at
            ship_date = parse_date_near_keywords(
                record.body,
                r"\b(?:Ship|Shipment|Shipping)\s*Date\b|出荷日|発送日|출하일|선적일",
                reference,
            ) or ship_date
            eta = parse_date_near_keywords(record.body, r"\bETA\b|到着予定|着港|도착예정", reference) or eta
            etd = parse_date_near_keywords(record.body, r"\bETD\b|出港|船積|출항", reference) or etd

        sales_by_entry = defaultdict(list)
        for event in events:
            if event.kind == "sales_note":
                sales_by_entry[event.entry_id].append(event)
        ledger: dict[tuple[str, str], BackorderSnapshot] = {}
        for record in records:
            event_snapshot = parse_backorders(
                record.body, mkr, record.received_at, f"메일 본문 {record.received_at:%Y-%m-%d %H:%M}"
            )
            seen_pdf_hashes: set[str] = set()
            for event in sales_by_entry[record.entry_id]:
                if event.sha256 in seen_pdf_hashes:
                    continue
                seen_pdf_hashes.add(event.sha256)
                event_snapshot.update(
                    parse_backorders(
                        read_pdf(event.path), mkr, event.received_at, f"SALES NOTE {event.original_name}"
                    )
                )
            merge_latest_backorders(ledger, event_snapshot, mkr)

        backorders = []
        for item in sorted(ledger.values(), key=lambda value: (value.source_mkr, value.code)):
            current = order_state.items.get(item.code)
            backorders.append(replace(item, origin=current.origin if current else "Unknown"))
        validate_capacity(mkr, len(order_state.items), len(backorders), config.max_rows)
        order_source = order_state.latest_event.original_name
        sales_source = sales_event.original_name if sales_event else ""
        decisions = build_decisions(
            mkr,
            order_state.items,
            sales_evidence,
            backorders,
            sales_note_available=sales_event is not None,
            order_source=order_source,
            sales_note_source=sales_source,
        )
        if sales_event is None:
            warnings.append("유효한 SALES NOTE가 없어 Shipped/Shortage QTY를 공란으로 유지했습니다.")
        if any(item.record_status == "SALES_NOTE_ONLY" for item in decisions):
            warnings.append("주문서에는 없고 SALES NOTE에서만 확인된 품목이 있습니다.")
        if any(item.origin == "Unknown" for item in decisions):
            warnings.append("파일명만으로 원산지를 판정하지 못한 품목이 있습니다.")
        status = "READY_WITH_WARNING" if warnings else "READY"
        results.append(
            MkrResult(
                mkr=mkr,
                sheet=sheet_name,
                decisions=decisions,
                backorders=backorders,
                order_date=order_state.first_event.received_at.date(),
                ship_date=ship_date,
                eta=eta,
                etd=etd,
                status=status,
                warnings=warnings,
                order_source=order_source,
                sales_note_source=sales_source,
                order_state_as_of=order_state.latest_event.received_at,
            )
        )
        logger.info(
            "[%s] 판정 완료: 현재 주문 %d개 / 백오더 %d개 / %s",
            mkr,
            len(decisions),
            len(backorders),
            status,
        )
    if not results:
        raise DataValidationError("읽을 수 있는 주문 자료가 없어 갱신할 시트가 없습니다.")
    return results
