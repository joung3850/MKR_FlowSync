from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .models import MkrResult


REPORT_FIELDS = (
    "Record Type",
    "MKR",
    "Back Order NO.",
    "Code",
    "Name",
    "Raw Order QTY",
    "First Shipped QTY",
    "Sales Note Order QTY",
    "K Total",
    "L Total",
    "Rule",
    "Final Order QTY",
    "Final Shipped QTY",
    "Final Shortage QTY",
    "Shipped Fallback",
    "Shipped Adjusted",
    "Origin",
    "Record Status",
    "Order Source",
    "Sales Note Source",
    "Order State As Of",
    "Source",
)


def write_audit_report(
    root: Path,
    results: Iterable[MkrResult],
    dry_run: bool,
    prefix: str | None = None,
) -> Path:
    report_dir = root / "Reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    report_prefix = prefix or ("MKR_DRY_RUN" if dry_run else "MKR_AUDIT")
    path = report_dir / f"{report_prefix}_{datetime.now():%Y%m%d_%H%M%S}.csv"
    def value_or_blank(value):
        return "" if value is None else value

    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_FIELDS)
        writer.writeheader()
        for result in results:
            for item in result.decisions:
                writer.writerow(
                    {
                        "Record Type": "NEW_ORDER",
                        "MKR": result.mkr,
                        "Code": item.code,
                        "Name": item.name,
                        "Raw Order QTY": value_or_blank(item.raw_order_qty),
                        "First Shipped QTY": value_or_blank(item.first_shipped_qty),
                        "Sales Note Order QTY": "" if item.sales_note_order_qty is None else item.sales_note_order_qty,
                        "K Total": item.backorder_order_qty,
                        "L Total": item.backorder_shipped_qty,
                        "Rule": item.rule_id,
                        "Final Order QTY": value_or_blank(item.order_qty),
                        "Final Shipped QTY": value_or_blank(item.shipped_qty),
                        "Final Shortage QTY": value_or_blank(item.shortage_qty),
                        "Shipped Fallback": "Y" if item.shipped_fallback else "N",
                        "Shipped Adjusted": "Y" if item.shipped_adjusted else "N",
                        "Origin": item.origin,
                        "Record Status": item.record_status,
                        "Order Source": item.order_source or result.order_source,
                        "Sales Note Source": item.sales_note_source or result.sales_note_source,
                        "Order State As Of": (
                            result.order_state_as_of.isoformat() if result.order_state_as_of else ""
                        ),
                    }
                )
            for item in result.backorders:
                writer.writerow(
                    {
                        "Record Type": "BACKORDER",
                        "MKR": result.mkr,
                        "Back Order NO.": item.source_mkr,
                        "Code": item.code,
                        "Name": item.name,
                        "K Total": item.order_qty,
                        "L Total": item.shipped_qty,
                        "Final Order QTY": item.order_qty,
                        "Final Shipped QTY": item.shipped_qty,
                        "Final Shortage QTY": item.shortage_qty,
                        "Source": item.source,
                        "Origin": item.origin,
                        "Record Status": "BACKORDER",
                        "Order Source": result.order_source,
                        "Sales Note Source": result.sales_note_source,
                        "Order State As Of": (
                            result.order_state_as_of.isoformat() if result.order_state_as_of else ""
                        ),
                    }
                )
    return path
