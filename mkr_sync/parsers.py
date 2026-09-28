from __future__ import annotations

import re
from collections import OrderedDict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Iterable, Sequence

from .errors import DataValidationError
from .models import BackorderSnapshot, OrderItem, SalesNoteEvidence, ZERO


CONTROL_SPLIT = re.compile(r"[\r\n\a\v\f]+")
MKR_PATTERN = re.compile(
    r"(?i)(?<![A-Z0-9])MKR\s*(\d{1,3})\s*[/_-]\s*(\d{2})(?:\s*[/_-]\s*(\d+))?"
)
NUMBER_PATTERN = r"[+-]?\d[\d,]*(?:\.\d+)?"


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_spaces(value: object) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def split_control_lines(text: str) -> list[str]:
    return [normalize_spaces(part) for part in CONTROL_SPLIT.split(text or "") if normalize_spaces(part)]


def parse_decimal(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None
    text = clean_text(value).replace(",", "")
    if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def canonical_mkr(value: object) -> str | None:
    match = MKR_PATTERN.search(clean_text(value))
    if not match:
        return None
    suffix = f"-{match.group(3)}" if match.group(3) else ""
    return f"MKR{match.group(1)}/{match.group(2)}{suffix}"


def find_all_mkrs(value: object) -> list[str]:
    found: list[str] = []
    for match in MKR_PATTERN.finditer(clean_text(value)):
        suffix = f"-{match.group(3)}" if match.group(3) else ""
        item = f"MKR{match.group(1)}/{match.group(2)}{suffix}"
        if item not in found:
            found.append(item)
    return found


def find_target_mkrs(value: object, targets: Iterable[str]) -> list[str]:
    allowed = set(targets)
    return [item for item in find_all_mkrs(value) if item in allowed]


def find_current_target_mkrs(subject: object, body: object, targets: Iterable[str]) -> list[str]:
    """Resolve destinations without treating historical back-order MKRs as current."""
    allowed = set(targets)
    subject_mkrs = find_all_mkrs(subject)
    if subject_mkrs:
        return [item for item in subject_mkrs if item in allowed]
    return find_target_mkrs(body, allowed)


def validate_mkr_match(expected_mkr: str, discovered: Iterable[str], label: str) -> None:
    values = sorted(set(discovered))
    # An order workbook may also contain older MKR identifiers in its
    # back-order table.  It is valid as long as the requested MKR itself is
    # present; reject only files that mention other MKRs without the target.
    if values and expected_mkr not in values:
        raise DataValidationError(
            f"메일 제목과 {label}의 MKR이 다릅니다: 제목={expected_mkr} / {label}={', '.join(values)}"
        )


def current_message_body(body: object) -> str:
    text = clean_text(body)
    cut = len(text)
    for pattern in (
        r"(?im)^\s*-----Original Message-----",
        r"(?im)^\s*From\s*:",
        r"(?im)^\s*差出人\s*:",
        r"(?im)^\s*送信者\s*:",
        r"(?im)^\s*보낸 사람\s*:",
    ):
        match = re.search(pattern, text)
        if match and 20 < match.start() < cut:
            cut = match.start()
    return text[:cut].strip()


def is_revision(value: object) -> bool:
    return bool(re.search(r"(?i)(?:^|[\s_\-(])rev(?:ised|ision)?(?:[\s_\-).]|$)", clean_text(value)))


def attachment_kind(filename: str, subject: str = "") -> tuple[str | None, bool]:
    name = clean_text(filename)
    revision = is_revision(name) or is_revision(subject)
    if re.search(r"(?i)sales[ _-]*note", name) and name.lower().endswith(".pdf"):
        return "sales_note", revision
    if not re.search(r"(?i)\.(xls|xlsx|xlsm)$", name):
        return None, revision
    # Files produced after the original order often contain the MKR number in
    # their name as well.  They are evidence for shipping/back-order work, but
    # must never become the Order QTY baseline.
    if re.search(
        r"(?i)lot\s*(?:no\.?|number)|pallet\s*detail|container\s*(?:weight|detail)|"
        r"invoice|packing\s*list|trade\s*shipping|(?:^|[_\s-])tsi(?:[_\s-]|$)|"
        r"出荷日変更|선적일\s*변경|출하일\s*변경|"
        r"(?:shipment|shipping)\s*date\s*(?:change|update)",
        name,
    ):
        return None, revision
    if not re.search(r"(?i)(\(krw\)|order|mkr)", name):
        return None, revision
    return "order", revision


def get_code(value: object) -> str | None:
    text = clean_text(value)
    match = re.search(r"(?<!\d)(\d{6})(?!\d)", text)
    if match:
        return match.group(1)
    number = parse_decimal(value)
    if number is not None and number == number.to_integral_value() and Decimal("10000") <= number <= Decimal("999999"):
        return f"{int(number):06d}"
    return None


def _is_code_header(value: object) -> bool:
    return bool(
        re.search(
            r"(?i)^(product\s*)?(item\s*)?(code|no\.?|number)$|product\s*code|item\s*(code|no)|商品.?コード|商品.?ｺｰﾄﾞ|商品.?cd|品番|품번",
            normalize_spaces(value),
        )
    )


def _is_name_header(value: object) -> bool:
    return bool(re.search(r"(?i)product\s*name|item\s*name|goods\s*name|description|商品名|品名|품명|상품명|^name$", normalize_spaces(value)))


def _is_qty_header(value: object) -> bool:
    text = normalize_spaces(value)
    if re.search(r"(?i)price|amount|金額|単価|가격|금액", text):
        return False
    return bool(re.search(r"(?i)order\s*(qty|quantity)|q[.\s-]?ty|(^|\s)(qty|quantity)(\s|$)|注文数|発注数|受注数|数量|수량", text))


def _is_non_item(value: object) -> bool:
    return bool(re.search(r"(?i)^(code|name|description|qty|quantity|pce|pcs|ea)$|price|amount|total|currency|krw|thb|usd|数量|金額|単価|합계|금액", normalize_spaces(value)))


def _is_suspicious_fallback_name(value: object) -> bool:
    """Reject common logistics identifiers found in non-order spreadsheets."""
    text = normalize_spaces(value)
    return bool(
        re.fullmatch(r"(?i)[A-Z]{4}\d{7}", text)
        or re.fullmatch(r"(?i)\d+(?:\.\d+)?\s*(?:FT|FOOT|FEET)(?:['\"])?", text)
    )


ORDER_ROW_BACKORDER_MARKER = re.compile(r"(?i)バック\s*オーダー|Back\s*[- ]?\s*Order")


def _add_order_item(
    items: dict[str, OrderItem],
    code: str,
    name: str,
    quantity: Decimal,
    *,
    backorders_separated: bool = False,
) -> None:
    if quantity < ZERO or not code or not name:
        return
    existing = items.get(code)
    if existing:
        items[code] = OrderItem(
            code,
            existing.name or name,
            existing.quantity + quantity,
            existing.backorders_separated or backorders_separated,
        )
    else:
        items[code] = OrderItem(code, name, quantity, backorders_separated)


def parse_order_matrix(matrix: Sequence[Sequence[object]]) -> tuple[dict[str, OrderItem], bool]:
    """Parse one Excel UsedRange matrix. Returns items and fallback-used flag."""
    rows = [list(row[:80]) for row in matrix[:1000]]
    items: dict[str, OrderItem] = OrderedDict()
    header: tuple[int, int, int, int] | None = None
    for row_index, row in enumerate(rows[:80]):
        code_col = name_col = qty_col = None
        for col_index, value in enumerate(row):
            if code_col is None and _is_code_header(value):
                code_col = col_index
            if name_col is None and _is_name_header(value):
                name_col = col_index
            if qty_col is None and _is_qty_header(value):
                qty_col = col_index
        if code_col is not None and name_col is not None and qty_col is not None:
            header = row_index, code_col, name_col, qty_col
            break

    if header:
        header_row, code_col, name_col, qty_col = header
        blank_run = 0
        filtered_backorder_rows = False
        for row in rows[header_row + 1 :]:
            code = get_code(row[code_col] if code_col < len(row) else None)
            if not code:
                blank_run += 1
                if blank_run >= 30 and items:
                    break
                continue
            blank_run = 0
            # The Milbon order workbook appends historical backorders to the
            # current MKR order and identifies them in the remarks column
            # (for example: "MKR20/26-1のバックオーダー").  These rows belong
            # in the separate backorder table and must never be summed into
            # the current MKR Order QTY.
            if ORDER_ROW_BACKORDER_MARKER.search(" ".join(clean_text(value) for value in row)):
                filtered_backorder_rows = True
                continue
            name = clean_text(row[name_col] if name_col < len(row) else None)
            qty = parse_decimal(row[qty_col] if qty_col < len(row) else None)
            if name and qty is not None:
                _add_order_item(items, code, name, qty)
        if filtered_backorder_rows:
            items = OrderedDict(
                (
                    code,
                    OrderItem(code, item.name, item.quantity, True),
                )
                for code, item in items.items()
            )
        return items, False

    for row in rows:
        code = None
        code_col = -1
        for col_index, value in enumerate(row):
            code = get_code(value)
            if code:
                code_col = col_index
                break
        if not code:
            continue
        name = ""
        quantity = None
        for raw in row[code_col + 1 :]:
            text = clean_text(raw)
            if not text:
                continue
            number = parse_decimal(raw)
            if (
                not name
                and number is None
                and not _is_non_item(text)
                and not _is_suspicious_fallback_name(text)
            ):
                name = text
                continue
            if name and number is not None and ZERO <= number <= Decimal("100000000"):
                quantity = number
                break
        if name and quantity is not None:
            _add_order_item(items, code, name, quantity)
    return items, True


BACKORDER_MARKER = re.compile(r"(?i)バック\s*オーダー|Back\s*[- ]?\s*Order(?:\s*QTY)?|欠品時\s*(?:No\.?|番号)")
BACKORDER_ROW = re.compile(
    rf"(?i)^\s*(?P<source>MKR\s*\d{{1,3}}\s*[/_-]\s*\d{{2}}(?:\s*[/_-]\s*\d+)?)\s+"
    rf"(?P<code>\d{{6}})\s+(?P<name>.+?)\s+(?P<order>{NUMBER_PATTERN})\s+"
    rf"(?P<shipped>{NUMBER_PATTERN})(?:\s+(?P<remaining>[-－―—]|{NUMBER_PATTERN}))?\s*$"
)


def text_before_backorders(text: str) -> str:
    match = BACKORDER_MARKER.search(text or "")
    return (text or "")[: match.start()].strip() if match else (text or "").strip()


def _parse_backorder_row(text: str, current_mkr: str, received_at: datetime, source: str) -> BackorderSnapshot | None:
    match = BACKORDER_ROW.match(normalize_spaces(text))
    if not match:
        return None
    source_mkr = canonical_mkr(match.group("source"))
    order_qty = parse_decimal(match.group("order"))
    shipped_qty = parse_decimal(match.group("shipped"))
    if not source_mkr or order_qty is None or shipped_qty is None or order_qty < ZERO or shipped_qty < ZERO:
        return None
    return BackorderSnapshot(
        current_mkr=current_mkr,
        source_mkr=source_mkr,
        code=match.group("code"),
        name=normalize_spaces(match.group("name")),
        order_qty=order_qty,
        shipped_qty=shipped_qty,
        received_at=received_at,
        source=source,
    )


def _parse_backorder_cells(
    lines: Sequence[str],
    index: int,
    current_mkr: str,
    received_at: datetime,
    source: str,
) -> tuple[BackorderSnapshot | None, int]:
    """Parse Outlook/Word tables whose cells are separated by control characters."""
    source_mkr = canonical_mkr(lines[index])
    if source_mkr is None or not re.fullmatch(MKR_PATTERN, lines[index]):
        return None, index
    if index + 3 >= len(lines) or not re.fullmatch(r"\d{6}", lines[index + 1]):
        return None, index

    code = lines[index + 1]
    name = normalize_spaces(lines[index + 2])
    quantities: list[Decimal] = []
    cursor = index + 3
    while cursor < len(lines):
        if re.match(r"(?i)^\s*MKR\s*\d{1,3}\s*[/_-]\s*\d{2}", lines[cursor]):
            break
        number = parse_decimal(lines[cursor])
        if number is not None:
            quantities.append(number)
        cursor += 1

    if not name or len(quantities) < 2:
        return None, index
    order_qty, shipped_qty = quantities[:2]
    if order_qty < ZERO or shipped_qty < ZERO:
        return None, index
    return (
        BackorderSnapshot(
            current_mkr=current_mkr,
            source_mkr=source_mkr,
            code=code,
            name=name,
            order_qty=order_qty,
            shipped_qty=shipped_qty,
            received_at=received_at,
            source=source,
        ),
        max(index, cursor - 1),
    )


def parse_backorders(text: str, current_mkr: str, received_at: datetime, source: str) -> dict[tuple[str, str], BackorderSnapshot]:
    marker = BACKORDER_MARKER.search(text or "")
    if not marker:
        return {}
    lines = split_control_lines((text or "")[marker.start() :])
    aggregated: dict[tuple[str, str], BackorderSnapshot] = OrderedDict()
    index = 0
    while index < len(lines):
        if not re.match(r"(?i)^\s*MKR\s*\d{1,3}\s*[/_-]\s*\d{2}", lines[index]):
            index += 1
            continue
        item = _parse_backorder_row(lines[index], current_mkr, received_at, source)
        consumed = index
        if item is None:
            item, consumed = _parse_backorder_cells(
                lines, index, current_mkr, received_at, source
            )
        if item is None:
            parts = [lines[index]]
            for cursor in range(index + 1, min(len(lines), index + 12)):
                if re.match(r"(?i)^\s*MKR\s*\d{1,3}\s*[/_-]\s*\d{2}", lines[cursor]):
                    break
                parts.append(lines[cursor])
                item = _parse_backorder_row(" ".join(parts), current_mkr, received_at, source)
                consumed = cursor
                if item is not None:
                    break
        if item is not None:
            # Repeated source-MKR + code rows in one Outlook body are usually
            # quoted reply-chain copies.  V11.8 keeps the last snapshot instead
            # of summing them, which prevents K/L from being multiplied.
            aggregated[item.key] = item
            index = consumed + 1
        else:
            index += 1
    return aggregated


def _add_sales_evidence(
    result: dict[str, SalesNoteEvidence],
    code: str,
    name: str,
    order_qty: Decimal | None,
    shipped_qty: Decimal | None,
) -> None:
    """Aggregate every physical Sales Note row for one product code.

    A code can appear once for the current MKR and again for one or more
    historical backorders.  Keeping only the first row loses the evidence
    needed to allocate those quantities correctly, so preserve their totals.
    """
    current = result.get(code)
    if current is None:
        result[code] = SalesNoteEvidence(code, name, order_qty, shipped_qty)
        return

    combined_order = current.order_qty
    if order_qty is not None:
        combined_order = (combined_order or ZERO) + order_qty
    combined_shipped = current.shipped_qty
    if shipped_qty is not None:
        combined_shipped = (combined_shipped or ZERO) + shipped_qty
    result[code] = SalesNoteEvidence(
        code,
        current.name or name,
        combined_order,
        combined_shipped,
    )


def parse_sales_note(text: str) -> dict[str, SalesNoteEvidence]:
    text = text_before_backorders(text)
    lines = split_control_lines(text)
    result: dict[str, SalesNoteEvidence] = OrderedDict()
    for line in lines:
        cleaned = re.sub(r"(?i)\b(PCE|PCS|EA)\b", " ", line)
        match = re.match(r"^\s*(\d{6})\s+(.+)$", cleaned)
        if not match:
            continue
        code, tail = match.group(1), match.group(2).strip()
        name_tokens: list[str] = []
        numbers: list[Decimal] = []
        numeric_started = False
        for token in tail.split():
            number = parse_decimal(token)
            if number is not None:
                numeric_started = True
                numbers.append(number)
            elif not numeric_started:
                name_tokens.append(token)
        if not name_tokens or not numbers:
            continue
        order_qty = numbers[0] if len(numbers) >= 2 else None
        shipped_qty = numbers[1] if len(numbers) >= 2 else numbers[0]
        _add_sales_evidence(
            result,
            code,
            " ".join(name_tokens),
            order_qty,
            shipped_qty,
        )

    for index, line in enumerate(lines):
        if not re.fullmatch(r"\d{6}", line):
            continue
        name = ""
        numbers: list[Decimal] = []
        for value in lines[index + 1 : index + 16]:
            if re.fullmatch(r"\d{6}", value):
                break
            if _is_code_header(value) or _is_name_header(value) or _is_qty_header(value):
                continue
            # ActiveReports Sales Notes place the code on one line and the
            # description plus shipped quantity on the next line, followed
            # by unit price and amount.  Read only the number immediately
            # before PCS/PCE/EA; currency amounts are not quantities.
            quantity_match = re.search(
                rf"(?i)(?P<qty>{NUMBER_PATTERN})\s*(?:PCE|PCS|EA)\b",
                value,
            )
            if quantity_match:
                description = value[: quantity_match.start()].strip()
                shipped_qty = parse_decimal(quantity_match.group("qty"))
                if description and shipped_qty is not None:
                    _add_sales_evidence(
                        result, line, description, None, shipped_qty
                    )
                break
            number = parse_decimal(value)
            if not name and number is None and not _is_non_item(value):
                name = value
            elif name and number is not None:
                numbers.append(number)
        if name and numbers:
            order_qty = numbers[0] if len(numbers) >= 2 else None
            shipped_qty = numbers[1] if len(numbers) >= 2 else numbers[0]
            _add_sales_evidence(result, line, name, order_qty, shipped_qty)
    return result


def merge_sales_evidence(
    primary: dict[str, SalesNoteEvidence], fallback: dict[str, SalesNoteEvidence]
) -> dict[str, SalesNoteEvidence]:
    merged = dict(primary)
    for code, extra in fallback.items():
        current = merged.get(code)
        if current is None:
            merged[code] = extra
            continue
        merged[code] = SalesNoteEvidence(
            code=code,
            name=current.name or extra.name,
            order_qty=current.order_qty if current.order_qty is not None else extra.order_qty,
            shipped_qty=current.shipped_qty if current.shipped_qty is not None else extra.shipped_qty,
        )
    return merged


def parse_date_near_keywords(body: str, keyword_pattern: str, reference: datetime) -> date | None:
    keyword = re.compile(keyword_pattern, re.IGNORECASE)
    month_map = {name: idx for idx, name in enumerate("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(), 1)}
    for raw_line in re.split(r"\r?\n", body or ""):
        line = raw_line.strip()
        match_keyword = keyword.search(line)
        if not match_keyword:
            continue
        candidate = line[match_keyword.end() :]
        matches = list(re.finditer(r"(?<!\d)(20\d{2})\s*[./\-年]\s*(\d{1,2})\s*[./\-月]\s*(\d{1,2})(?:일)?", candidate))
        if matches:
            m = matches[-1]
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
        matches = list(re.finditer(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(20\d{2})(?!\d)", candidate))
        if matches:
            m = matches[-1]
            try:
                return date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
            except ValueError:
                pass
        english = re.search(r"(?i)(\d{1,2})[- ](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[- ,](20\d{2})", candidate)
        if english:
            try:
                return date(int(english.group(3)), month_map[english.group(2).title()[:3]], int(english.group(1)))
            except ValueError:
                pass
        english = re.search(
            r"(?i)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(20\d{2})",
            candidate,
        )
        if english:
            try:
                return date(int(english.group(3)), month_map[english.group(1).title()[:3]], int(english.group(2)))
            except ValueError:
                pass
        fallback = list(re.finditer(r"(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)", candidate))
        if fallback:
            m = fallback[-1]
            try:
                return date(reference.year, int(m.group(1)), int(m.group(2)))
            except ValueError:
                pass
    return None
