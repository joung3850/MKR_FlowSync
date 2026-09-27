from __future__ import annotations

import re


# Deliberately does not match P/O identifiers such as MKR-20260910TG-A(T).
MKR_PATTERN = re.compile(r"MKR\s*(\d+)\s*/\s*(\d+)(?:\s*-\s*(\d+))?", re.IGNORECASE)
EXCEL_FORBIDDEN_SHEET_CHARS = set(':\\/?*[]')


def extract_mkr_numbers(text: str) -> list[str]:
    """Extract unique MKR shipment identifiers in first-seen order."""
    found: list[str] = []
    seen: set[str] = set()
    for match in MKR_PATTERN.finditer(text or ""):
        suffix = f"-{match.group(3)}" if match.group(3) else ""
        value = f"MKR{match.group(1)}/{match.group(2)}{suffix}".upper()
        if value not in seen:
            seen.add(value)
            found.append(value)
    return found


def to_excel_sheet_name(mkr_number: str) -> str:
    """Convert a normalized MKR number to an Excel-compatible sheet name."""
    value = mkr_number.strip().upper().replace("/", "_")
    if not value:
        raise ValueError("MKR 번호가 비어 있습니다.")
    if len(value) > 31:
        raise ValueError(f"Excel 시트 이름은 31자를 초과할 수 없습니다: {value}")
    if any(char in value for char in EXCEL_FORBIDDEN_SHEET_CHARS):
        raise ValueError(f"Excel 시트 이름에 사용할 수 없는 문자가 있습니다: {value}")
    return value
