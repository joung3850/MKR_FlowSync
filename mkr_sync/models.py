from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path


ZERO = Decimal("0")


@dataclass(frozen=True)
class OrderItem:
    code: str
    name: str
    quantity: Decimal


@dataclass(frozen=True)
class SalesNoteEvidence:
    code: str
    name: str
    order_qty: Decimal | None
    shipped_qty: Decimal | None


@dataclass(frozen=True)
class BackorderSnapshot:
    current_mkr: str
    source_mkr: str
    code: str
    name: str
    order_qty: Decimal
    shipped_qty: Decimal
    received_at: datetime
    source: str

    @property
    def shortage_qty(self) -> Decimal:
        return max(ZERO, self.order_qty - self.shipped_qty)

    @property
    def key(self) -> tuple[str, str]:
        return self.source_mkr, self.code


@dataclass(frozen=True)
class QuantityDecision:
    mkr: str
    code: str
    name: str
    raw_order_qty: Decimal
    first_shipped_qty: Decimal
    sales_note_order_qty: Decimal | None
    backorder_order_qty: Decimal
    backorder_shipped_qty: Decimal
    order_qty: Decimal
    shipped_qty: Decimal
    rule_id: str
    shipped_fallback: bool = False
    shipped_adjusted: bool = False

    @property
    def shortage_qty(self) -> Decimal:
        return max(ZERO, self.order_qty - self.shipped_qty)


@dataclass
class MkrResult:
    mkr: str
    sheet: str
    decisions: list[QuantityDecision]
    backorders: list[BackorderSnapshot]
    order_date: date | None = None
    ship_date: date | None = None
    eta: date | None = None
    etd: date | None = None


@dataclass(frozen=True)
class MailRecord:
    entry_id: str
    target_mkr: str
    received_at: datetime
    subject: str
    body: str


@dataclass(frozen=True)
class AttachmentEvent:
    entry_id: str
    target_mkr: str
    received_at: datetime
    subject: str
    body: str
    kind: str
    is_revision: bool
    original_name: str
    path: Path
    sha256: str


@dataclass
class CollectedMail:
    records: list[MailRecord] = field(default_factory=list)
    attachments: list[AttachmentEvent] = field(default_factory=list)
    scanned: int = 0
