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
    "Source",
)


def write_audit_report(root: Path, results: Iterable[MkrResult], dry_run: bool) -> Path:
    report_dir = root / "Reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    prefix = "MKR_DRY_RUN" if dry_run else "MKR_AUDIT"
    path = report_dir / f"{prefix}_{datetime.now():%Y%m%d_%H%M%S}.csv"
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
                        "Raw Order QTY": item.raw_order_qty,
                        "First Shipped QTY": item.first_shipped_qty,
                        "Sales Note Order QTY": "" if item.sales_note_order_qty is None else item.sales_note_order_qty,
                        "K Total": item.backorder_order_qty,
                        "L Total": item.backorder_shipped_qty,
                        "Rule": item.rule_id,
                        "Final Order QTY": item.order_qty,
                        "Final Shipped QTY": item.shipped_qty,
                        "Final Shortage QTY": item.shortage_qty,
                        "Shipped Fallback": "Y" if item.shipped_fallback else "N",
                        "Shipped Adjusted": "Y" if item.shipped_adjusted else "N",
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
                    }
                )
    return path
