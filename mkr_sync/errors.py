from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal


class MkrSyncError(Exception):
    """Base error for a safe, user-visible synchronization stop."""


class ConfigurationError(MkrSyncError):
    pass


class OfficeAutomationError(MkrSyncError):
    pass


class DataValidationError(MkrSyncError):
    pass


class WorkbookCommitError(MkrSyncError):
    pass


@dataclass(frozen=True)
class RuleContext:
    mkr: str
    code: str
    raw_order_qty: Decimal
    first_shipped_qty: Decimal
    sales_note_order_qty: Decimal | None
    backorder_order_qty: Decimal
    backorder_shipped_qty: Decimal
    attempted_rule: str

    def as_log_text(self) -> str:
        p = "누락" if self.sales_note_order_qty is None else str(self.sales_note_order_qty)
        return (
            f"MKR={self.mkr} / 품번={self.code} / 최초 Order={self.raw_order_qty} / "
            f"최초 Shipped={self.first_shipped_qty} / Sales Note Order={p} / "
            f"K={self.backorder_order_qty} / L={self.backorder_shipped_qty} / "
            f"시도 규칙={self.attempted_rule}"
        )


class QuantityRuleError(DataValidationError):
    def __init__(self, message: str, context: RuleContext):
        self.context = context
        super().__init__(f"{message} | {context.as_log_text()}")
