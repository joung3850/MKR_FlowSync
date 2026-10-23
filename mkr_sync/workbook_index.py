from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Iterable, Literal


LOOKUP_SHEET = "MKR조회"
SUMMARY_SHEET = "품목합계"
MANAGED_SHEETS = frozenset({LOOKUP_SHEET, SUMMARY_SHEET})
LOOKUP_TABLE = "MkrLookupTable"
SUMMARY_TABLE = "ProductSummaryTable"

LOOKUP_HEADERS = (
    "MKR",
    "품번",
    "품명",
    "원산지",
    "신규 Order",
    "신규 Shipped",
    "신규 Shortage",
    "백오더 Order",
    "백오더 Shipped",
    "백오더 Shortage",
    "총 수입 Shipped",
    "데이터 상태",
)

SUMMARY_HEADERS = (
    "품번",
    "품명",
    "원산지",
    "포함 MKR 수",
    "신규 Order 합계",
    "신규 Shipped 합계",
    "백오더 Shipped 합계",
    "총 수입 Shipped",
    "신규 Shortage 합계",
    "최신 백오더 Order",
    "최신 백오더 Shortage",
    "데이터 상태",
)

MKR_SHEET_PATTERN = re.compile(r"^MKR(\d+)_(\d+)(?:-(\d+))?$", re.IGNORECASE)


@dataclass(frozen=True)
class WorkbookRecord:
    current_mkr: str
    code: str
    name: str
    origin: str
    record_type: Literal["CURRENT", "BACKORDER"]
    order_qty: Decimal | None
    shipped_qty: Decimal | None
    shortage_qty: Decimal | None
    backorder_no: str = ""
    row_number: int = 0

    @property
    def sort_key(self) -> tuple[int, int, int, int]:
        year, number, suffix = mkr_sort_key(self.current_mkr)
        return year, number, suffix, self.row_number


@dataclass(frozen=True)
class MkrLookupRow:
    mkr: str
    code: str
    name: str
    origin: str
    new_order_qty: Decimal | None
    new_shipped_qty: Decimal | None
    new_shortage_qty: Decimal | None
    backorder_order_qty: Decimal | None
    backorder_shipped_qty: Decimal | None
    backorder_shortage_qty: Decimal | None
    total_imported_shipped_qty: Decimal | None
    data_status: str

    def excel_row(self) -> tuple[object, ...]:
        return (
            self.mkr,
            self.code,
            self.name,
            self.origin,
            self.new_order_qty,
            self.new_shipped_qty,
            self.new_shortage_qty,
            self.backorder_order_qty,
            self.backorder_shipped_qty,
            self.backorder_shortage_qty,
            self.total_imported_shipped_qty,
            self.data_status,
        )


@dataclass(frozen=True)
class ProductSummaryRow:
    code: str
    name: str
    origin: str
    mkr_count: int
    new_order_qty: Decimal | None
    new_shipped_qty: Decimal | None
    backorder_shipped_qty: Decimal | None
    total_imported_shipped_qty: Decimal | None
    new_shortage_qty: Decimal | None
    latest_backorder_order_qty: Decimal | None
    latest_backorder_shortage_qty: Decimal | None
    data_status: str

    def excel_row(self) -> tuple[object, ...]:
        return (
            self.code,
            self.name,
            self.origin,
            self.mkr_count,
            self.new_order_qty,
            self.new_shipped_qty,
            self.backorder_shipped_qty,
            self.total_imported_shipped_qty,
            self.new_shortage_qty,
            self.latest_backorder_order_qty,
            self.latest_backorder_shortage_qty,
            self.data_status,
        )


@dataclass(frozen=True)
class WorkbookIndex:
    lookup_rows: tuple[MkrLookupRow, ...]
    product_rows: tuple[ProductSummaryRow, ...]


def mkr_sort_key(mkr: str) -> tuple[int, int, int]:
    match = re.fullmatch(r"MKR(\d+)/(\d+)(?:-(\d+))?", str(mkr).strip(), re.IGNORECASE)
    if match is None:
        return 0, 0, 0
    return int(match.group(2)), int(match.group(1)), int(match.group(3) or 0)


def mkr_from_sheet_name(sheet_name: str) -> str | None:
    match = MKR_SHEET_PATTERN.fullmatch(str(sheet_name).strip())
    if match is None:
        return None
    suffix = f"-{int(match.group(3))}" if match.group(3) else ""
    return f"MKR{int(match.group(1))}/{int(match.group(2))}{suffix}"


def normalize_code(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        decimal_value = Decimal(text)
    except InvalidOperation:
        return text
    if decimal_value == decimal_value.to_integral_value():
        digits = str(int(decimal_value))
        return digits.zfill(6) if digits.isdigit() and len(digits) <= 6 else digits
    return text


def optional_decimal(value: object) -> Decimal | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return None


def _sum_known(records: Iterable[WorkbookRecord], field: str) -> Decimal | None:
    values = [getattr(item, field) for item in records]
    if not values:
        return Decimal("0")
    known = [value for value in values if value is not None]
    return sum(known, Decimal("0")) if known else None


def _latest_name(records: Iterable[WorkbookRecord]) -> str:
    ordered = sorted(records, key=lambda item: item.sort_key)
    return next((item.name.strip() for item in reversed(ordered) if item.name.strip()), "")


def _origins(records: Iterable[WorkbookRecord]) -> str:
    values: set[str] = set()
    for item in records:
        value = item.origin.strip() or "Unknown"
        if value == "Mixed":
            values.update(("Japan", "Thailand"))
        else:
            values.add(value)
    if len(values) > 1:
        values.discard("Unknown")
    order = {"Japan": 0, "Thailand": 1, "Unknown": 9}
    return ", ".join(sorted(values or {"Unknown"}, key=lambda value: (order.get(value, 5), value)))


def _status(groups: Iterable[tuple[str, Iterable[WorkbookRecord], str]]) -> str:
    messages: list[str] = []
    for label, records, field in groups:
        missing = sum(getattr(item, field) is None for item in records)
        if missing:
            messages.append(f"{label} {missing}건 미확인")
    return ", ".join(messages) if messages else "완료"


def _dedupe_backorders(records: Iterable[WorkbookRecord]) -> list[WorkbookRecord]:
    latest: dict[tuple[str, str, str], WorkbookRecord] = {}
    for item in records:
        key = (item.current_mkr, item.backorder_no, item.code)
        previous = latest.get(key)
        if previous is None or item.sort_key >= previous.sort_key:
            latest[key] = item
    return list(latest.values())


def build_workbook_index(records: Iterable[WorkbookRecord]) -> WorkbookIndex:
    source = [item for item in records if item.code]
    deduped_backorders = _dedupe_backorders(
        item for item in source if item.record_type == "BACKORDER"
    )

    lookup_groups: dict[tuple[str, str], list[WorkbookRecord]] = defaultdict(list)
    for item in source:
        if item.record_type == "CURRENT":
            lookup_groups[(item.current_mkr, item.code)].append(item)
    for item in deduped_backorders:
        lookup_groups[(item.current_mkr, item.code)].append(item)

    lookup_rows: list[MkrLookupRow] = []
    for (mkr, code), group in lookup_groups.items():
        current = [item for item in group if item.record_type == "CURRENT"]
        backorders = [item for item in group if item.record_type == "BACKORDER"]
        new_shipped = _sum_known(current, "shipped_qty")
        backorder_shipped = _sum_known(backorders, "shipped_qty")
        lookup_rows.append(
            MkrLookupRow(
                mkr=mkr,
                code=code,
                name=_latest_name(group),
                origin=_origins(group),
                new_order_qty=_sum_known(current, "order_qty"),
                new_shipped_qty=new_shipped,
                new_shortage_qty=_sum_known(current, "shortage_qty"),
                backorder_order_qty=_sum_known(backorders, "order_qty"),
                backorder_shipped_qty=backorder_shipped,
                backorder_shortage_qty=_sum_known(backorders, "shortage_qty"),
                total_imported_shipped_qty=_sum_known(group, "shipped_qty"),
                data_status=_status(
                    (
                        ("신규 Order", current, "order_qty"),
                        ("신규 Shipped", current, "shipped_qty"),
                        ("신규 Shortage", current, "shortage_qty"),
                        ("백오더 Order", backorders, "order_qty"),
                        ("백오더 Shipped", backorders, "shipped_qty"),
                        ("백오더 Shortage", backorders, "shortage_qty"),
                    )
                ),
            )
        )
    lookup_rows.sort(key=lambda item: (*mkr_sort_key(item.mkr), item.code))

    latest_backorders: dict[tuple[str, str], WorkbookRecord] = {}
    for item in deduped_backorders:
        key = (item.backorder_no, item.code)
        previous = latest_backorders.get(key)
        if previous is None or item.sort_key >= previous.sort_key:
            latest_backorders[key] = item

    product_groups: dict[str, list[WorkbookRecord]] = defaultdict(list)
    for item in source:
        if item.record_type == "CURRENT":
            product_groups[item.code].append(item)
    for item in deduped_backorders:
        product_groups[item.code].append(item)

    product_rows: list[ProductSummaryRow] = []
    for code, group in product_groups.items():
        current = [item for item in group if item.record_type == "CURRENT"]
        shipped_backorders = [item for item in group if item.record_type == "BACKORDER"]
        latest_for_code = [
            item for (source_mkr, item_code), item in latest_backorders.items() if item_code == code
        ]
        new_shipped = _sum_known(current, "shipped_qty")
        backorder_shipped = _sum_known(shipped_backorders, "shipped_qty")
        product_rows.append(
            ProductSummaryRow(
                code=code,
                name=_latest_name(group),
                origin=_origins(group),
                mkr_count=len({item.current_mkr for item in group}),
                new_order_qty=_sum_known(current, "order_qty"),
                new_shipped_qty=new_shipped,
                backorder_shipped_qty=backorder_shipped,
                total_imported_shipped_qty=_sum_known(group, "shipped_qty"),
                new_shortage_qty=_sum_known(current, "shortage_qty"),
                latest_backorder_order_qty=_sum_known(latest_for_code, "order_qty"),
                latest_backorder_shortage_qty=_sum_known(latest_for_code, "shortage_qty"),
                data_status=_status(
                    (
                        ("신규 Order", current, "order_qty"),
                        ("신규 Shipped", current, "shipped_qty"),
                        ("신규 Shortage", current, "shortage_qty"),
                        ("백오더 Shipped", shipped_backorders, "shipped_qty"),
                        ("최신 백오더 Order", latest_for_code, "order_qty"),
                        ("최신 백오더 Shortage", latest_for_code, "shortage_qty"),
                    )
                ),
            )
        )
    product_rows.sort(key=lambda item: item.code)
    return WorkbookIndex(tuple(lookup_rows), tuple(product_rows))


def _as_rows(value: object) -> list[list[object]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def extract_workbook_records(workbook, logger, first_row: int = 10, last_row: int = 509) -> list[WorkbookRecord]:
    records: list[WorkbookRecord] = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets.Item(index)
        sheet_name = str(sheet.Name)
        expected_mkr = mkr_from_sheet_name(sheet_name)
        if expected_mkr is None:
            continue
        actual_mkr = str(sheet.Range("D3").Value2 or "").strip().upper().replace(" ", "")
        if actual_mkr != expected_mkr.upper():
            logger.warning(
                "통합조회 제외: %s의 D3 MKR 값이 일치하지 않습니다 (%s)",
                sheet_name,
                actual_mkr or "공란",
            )
            continue

        current_rows = _as_rows(sheet.Range(f"A{first_row}:G{last_row}").Value2)
        for offset, row in enumerate(current_rows):
            code = normalize_code(row[0] if len(row) > 0 else None)
            if not code:
                continue
            records.append(
                WorkbookRecord(
                    current_mkr=expected_mkr,
                    code=code,
                    name=str(row[1] or "").strip(),
                    order_qty=optional_decimal(row[2]),
                    shipped_qty=optional_decimal(row[3]),
                    shortage_qty=optional_decimal(row[4]),
                    origin=str(row[6] or "Unknown").strip() or "Unknown",
                    record_type="CURRENT",
                    row_number=first_row + offset,
                )
            )

        backorder_rows = _as_rows(sheet.Range(f"H{first_row}:N{last_row}").Value2)
        for offset, row in enumerate(backorder_rows):
            code = normalize_code(row[1] if len(row) > 1 else None)
            if not code:
                continue
            records.append(
                WorkbookRecord(
                    current_mkr=expected_mkr,
                    backorder_no=str(row[0] or "").strip(),
                    code=code,
                    name=str(row[2] or "").strip(),
                    order_qty=optional_decimal(row[3]),
                    shipped_qty=optional_decimal(row[4]),
                    shortage_qty=optional_decimal(row[5]),
                    origin=str(row[6] or "Unknown").strip() or "Unknown",
                    record_type="BACKORDER",
                    row_number=first_row + offset,
                )
            )
    return records


def _excel_value(value: object) -> object:
    if not isinstance(value, Decimal):
        return value
    return int(value) if value == value.to_integral_value() else float(value)


def _write_managed_sheet(
    workbook,
    sheet,
    *,
    title: str,
    note: str,
    headers: tuple[str, ...],
    rows: Iterable[tuple[object, ...]],
    table_name: str,
    updated_at: datetime,
) -> None:
    materialized = [tuple(_excel_value(value) for value in row) for row in rows]
    last_column = chr(ord("A") + len(headers) - 1)
    sheet.Cells.Font.Name = "Arial"
    sheet.Cells.Font.Size = 10
    sheet.Range("A2").Value2 = title
    sheet.Range("A2").Font.Bold = True
    sheet.Range("A2").Font.Size = 14
    sheet.Range("A2").Font.Color = 0x432A10
    sheet.Range("A3").Value2 = f"마지막 갱신: {updated_at:%Y-%m-%d %H:%M:%S}"
    sheet.Range("A3").Font.Italic = True
    sheet.Range("A3").Font.Color = 0x7A6655
    sheet.Range("A4").Value2 = note
    sheet.Range("A4").Font.Color = 0x7A6655
    sheet.Range(f"A6:{last_column}6").Value2 = (headers,)
    sheet.Range(f"A6:{last_column}6").Font.Bold = True
    sheet.Range(f"A6:{last_column}6").Font.Color = 0xFFFFFF
    sheet.Range(f"A6:{last_column}6").Interior.Color = 0x432A10
    sheet.Range(f"A6:{last_column}6").HorizontalAlignment = -4108
    sheet.Range(f"A6:{last_column}6").VerticalAlignment = -4108

    data_last_row = 6 + max(len(materialized), 1)
    if sheet.Name == SUMMARY_SHEET:
        sheet.Range(f"A7:C{data_last_row}").NumberFormat = "@"
        sheet.Range(f"D7:D{data_last_row}").NumberFormat = "#,##0"
        sheet.Range(f"E7:K{data_last_row}").NumberFormat = "#,##0.###"
    else:
        sheet.Range(f"A7:D{data_last_row}").NumberFormat = "@"
        sheet.Range(f"E7:K{data_last_row}").NumberFormat = "#,##0.###"
    if materialized:
        sheet.Range(f"A7:{last_column}{data_last_row}").Value2 = tuple(materialized)
    else:
        sheet.Range(f"A7:{last_column}7").ClearContents()
    table = sheet.ListObjects.Add(1, sheet.Range(f"A6:{last_column}{data_last_row}"), None, 1)
    table.Name = table_name
    table.TableStyle = "TableStyleMedium2"
    table.ShowAutoFilter = True
    sheet.Range(f"A7:{last_column}{data_last_row}").VerticalAlignment = -4108
    sheet.Range(f"A7:{last_column}{data_last_row}").Borders.LineStyle = 0

    sheet.Columns("A").ColumnWidth = 15
    sheet.Columns("B").ColumnWidth = 13
    sheet.Columns("C").ColumnWidth = 42
    sheet.Columns("D").ColumnWidth = 18
    for column in range(5, 12):
        sheet.Columns(column).ColumnWidth = 17
    sheet.Columns(12).ColumnWidth = 32
    if sheet.Name == SUMMARY_SHEET:
        sheet.Columns("A").ColumnWidth = 13
        sheet.Columns("B").ColumnWidth = 42
        sheet.Columns("C").ColumnWidth = 18
        sheet.Columns("D").ColumnWidth = 13
    sheet.Activate()
    window = workbook.Application.ActiveWindow
    window.SplitRow = 6
    window.SplitColumn = 2
    window.FreezePanes = True
    window.DisplayGridlines = False


def rebuild_lookup_sheets(workbook, logger, updated_at: datetime | None = None) -> WorkbookIndex:
    records = extract_workbook_records(workbook, logger)
    index = build_workbook_index(records)
    existing = {
        str(workbook.Worksheets.Item(i).Name): workbook.Worksheets.Item(i)
        for i in range(1, int(workbook.Worksheets.Count) + 1)
    }
    for name in (LOOKUP_SHEET, SUMMARY_SHEET):
        sheet = existing.get(name)
        if sheet is not None:
            sheet.Delete()

    summary_sheet = workbook.Worksheets.Add(Before=workbook.Worksheets.Item(1))
    summary_sheet.Name = SUMMARY_SHEET
    lookup_sheet = workbook.Worksheets.Add(Before=workbook.Worksheets.Item(1))
    lookup_sheet.Name = LOOKUP_SHEET
    refreshed_at = updated_at or datetime.now()
    _write_managed_sheet(
        workbook,
        lookup_sheet,
        title="MKR별 품목 조회",
        note=(
            "표 필터에서 MKR·품번·품명을 검색합니다. 총 수입 Shipped는 신규 Shipped와 "
            "이번 수입건의 백오더 Shipped 합계입니다. 이 시트의 수동 수정은 다음 적용 시 재생성됩니다."
        ),
        headers=LOOKUP_HEADERS,
        rows=(item.excel_row() for item in index.lookup_rows),
        table_name=LOOKUP_TABLE,
        updated_at=refreshed_at,
    )
    _write_managed_sheet(
        workbook,
        summary_sheet,
        title="품목별 전체 수량",
        note=(
            "모든 MKR의 품목별 합계입니다. 백오더 Order·Shortage는 Back Order NO.별 최신 기록만 "
            "집계하며, Shipped는 각 수입건의 실적을 합산합니다."
        ),
        headers=SUMMARY_HEADERS,
        rows=(item.excel_row() for item in index.product_rows),
        table_name=SUMMARY_TABLE,
        updated_at=refreshed_at,
    )
    lookup_sheet.Activate()
    logger.info(
        "통합조회 시트 갱신: MKR별 품목 %d개 / 전체 품목 %d개",
        len(index.lookup_rows),
        len(index.product_rows),
    )
    return index


def _table_names(sheet) -> set[str]:
    return {
        str(sheet.ListObjects.Item(index).Name)
        for index in range(1, int(sheet.ListObjects.Count) + 1)
    }


def _compare_written_rows(sheet, expected_rows: tuple[tuple[object, ...], ...]) -> None:
    if not expected_rows:
        return
    actual_rows = _as_rows(sheet.Range(f"A7:L{6 + len(expected_rows)}").Value2)
    if len(actual_rows) != len(expected_rows):
        raise ValueError("조회 데이터 행 수가 다릅니다.")
    for row_index, (actual, expected) in enumerate(zip(actual_rows, expected_rows), start=7):
        for column_index, (actual_value, expected_value) in enumerate(zip(actual, expected), start=1):
            if isinstance(expected_value, Decimal):
                actual_decimal = optional_decimal(actual_value)
                if actual_decimal != expected_value:
                    raise ValueError(f"조회 데이터 수량이 다릅니다: {row_index}행 {column_index}열")
            elif actual_value != expected_value:
                raise ValueError(f"조회 데이터 값이 다릅니다: {row_index}행 {column_index}열")


def validate_lookup_sheets(workbook, index: WorkbookIndex) -> None:
    names = [str(workbook.Worksheets.Item(i).Name) for i in range(1, int(workbook.Worksheets.Count) + 1)]
    if names[:2] != [LOOKUP_SHEET, SUMMARY_SHEET]:
        raise ValueError("통합조회 시트 순서가 올바르지 않습니다.")
    lookup = workbook.Worksheets.Item(LOOKUP_SHEET)
    summary = workbook.Worksheets.Item(SUMMARY_SHEET)
    if tuple(lookup.Range("A6:L6").Value2[0]) != LOOKUP_HEADERS:
        raise ValueError("MKR조회 헤더가 올바르지 않습니다.")
    if tuple(summary.Range("A6:L6").Value2[0]) != SUMMARY_HEADERS:
        raise ValueError("품목합계 헤더가 올바르지 않습니다.")
    if LOOKUP_TABLE not in _table_names(lookup) or SUMMARY_TABLE not in _table_names(summary):
        raise ValueError("통합조회 Excel Table이 누락되었습니다.")

    lookup_expected = tuple(item.excel_row() for item in index.lookup_rows)
    summary_expected = tuple(item.excel_row() for item in index.product_rows)
    _compare_written_rows(lookup, lookup_expected)
    _compare_written_rows(summary, summary_expected)
    for sheet, row_count in ((lookup, len(lookup_expected)), (summary, len(summary_expected))):
        formulas = _as_rows(sheet.Range(f"A1:L{max(7, row_count + 6)}").Formula)
        if any(isinstance(value, str) and value.startswith("=") for row in formulas for value in row):
            raise ValueError(f"{sheet.Name}에 예상하지 않은 수식이 있습니다.")
