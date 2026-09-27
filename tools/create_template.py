from __future__ import annotations

import hashlib
from pathlib import Path

import pythoncom
import win32com.client


SOURCE_SHEETS = (
    "MKR56_26",
    "MKR48_26-1",
    "MKR48_26",
    "MKR49_26",
    "MKR49_26-1",
)
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
XL_CELL_TYPE_CONSTANTS = 2
XL_OPEN_XML_WORKBOOK = 51


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def clean(value: object) -> str:
    return "" if value is None else str(value).strip()


def merged_ranges(sheet) -> tuple[str, ...]:
    merged: set[str] = set()
    try:
        areas = sheet.UsedRange.MergeAreas
        for index in range(1, int(areas.Count) + 1):
            merged.add(str(areas.Item(index).Address))
    except Exception:
        pass
    return tuple(sorted(merged))


def fingerprint(sheet) -> tuple:
    used = sheet.UsedRange
    return (
        str(used.Address),
        int(used.Rows.Count),
        int(used.Columns.Count),
        merged_ranges(sheet),
        tuple(clean(value) for value in sheet.Range("A9:M9").Value2[0]),
    )


def clear_constants(sheet, address: str) -> None:
    target = sheet.Range(address)
    if int(target.Cells.Count) == 1:
        if not bool(target.HasFormula):
            target.ClearContents()
        return
    try:
        target.SpecialCells(XL_CELL_TYPE_CONSTANTS).ClearContents()
    except Exception:
        pass


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    source_path = root / "MKR_TEST.xlsx"
    output_path = root / "MKR_TEMPLATE.xlsx"
    source_hash = digest(source_path)
    pythoncom.CoInitialize()
    excel = source = template_book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        source = excel.Workbooks.Open(str(source_path), 0, True)
        names = {
            clean(source.Worksheets.Item(index).Name)
            for index in range(1, int(source.Worksheets.Count) + 1)
        }
        missing = [name for name in SOURCE_SHEETS if name not in names]
        if missing:
            raise RuntimeError("기준 시트 누락: " + ", ".join(missing))
        expected = fingerprint(source.Worksheets.Item(SOURCE_SHEETS[0]))
        mismatched = [
            name
            for name in SOURCE_SHEETS[1:]
            if fingerprint(source.Worksheets.Item(name)) != expected
        ]
        if mismatched:
            raise RuntimeError("기준 시트 구조 불일치: " + ", ".join(mismatched))

        source.Worksheets.Item(SOURCE_SHEETS[0]).Copy()
        template_book = excel.ActiveWorkbook
        copied = template_book.Worksheets.Item(1)
        copied.Name = "MKR_TEMPLATE"
        for address in ("A10:E509", "H10:M509", "D3", "K4:K7"):
            clear_constants(copied, address)
        template_book.SaveAs(str(output_path), FileFormat=XL_OPEN_XML_WORKBOOK)
        template_book.Close(False)
        template_book = None

        check = excel.Workbooks.Open(str(output_path), 0, True)
        try:
            if int(check.Worksheets.Count) != 1 or clean(check.Worksheets.Item(1).Name) != "MKR_TEMPLATE":
                raise RuntimeError("템플릿은 MKR_TEMPLATE 시트 하나만 포함해야 합니다.")
            sheet = check.Worksheets.Item("MKR_TEMPLATE")
            headers = tuple(clean(value) for value in sheet.Range("A9:M9").Value2[0])
            for index, expected_header in EXPECTED_HEADERS.items():
                if headers[index] != expected_header:
                    raise RuntimeError(f"템플릿 헤더 오류: {index + 1}열")
            for address in ("A10:E509", "H10:M509", "D3", "K4:K7"):
                target = sheet.Range(address)
                values = target.Value2
                formulas = target.Formula
                value_rows = values if isinstance(values, tuple) else ((values,),)
                formula_rows = formulas if isinstance(formulas, tuple) else ((formulas,),)
                for value_row, formula_row in zip(value_rows, formula_rows):
                    value_cells = value_row if isinstance(value_row, tuple) else (value_row,)
                    formula_cells = formula_row if isinstance(formula_row, tuple) else (formula_row,)
                    for value, formula in zip(value_cells, formula_cells):
                        is_formula = isinstance(formula, str) and formula.startswith("=")
                        if not is_formula and value not in (None, ""):
                            raise RuntimeError(f"템플릿 데이터가 남아 있습니다: {address}")
            if fingerprint(sheet) != expected:
                raise RuntimeError("템플릿 구조가 기준 시트와 다릅니다.")
        finally:
            check.Close(False)
    finally:
        if template_book is not None:
            template_book.Close(False)
        if source is not None:
            source.Close(False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    if digest(source_path) != source_hash:
        raise RuntimeError("원본 MKR_TEST.xlsx가 변경되었습니다.")
    print(output_path)
    print("source_sha256=" + source_hash)
    print("template_sha256=" + digest(output_path))


if __name__ == "__main__":
    main()
