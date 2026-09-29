from __future__ import annotations

from collections import defaultdict
from dataclasses import replace
from decimal import Decimal
from typing import Iterable

from .errors import QuantityRuleError, RuleContext
from .models import BackorderSnapshot, OrderItem, QuantityDecision, SalesNoteEvidence, ZERO


def _ensure_nonnegative(value: Decimal, label: str, context: RuleContext) -> None:
    if value < ZERO:
        raise QuantityRuleError(f"{label}가 음수입니다.", context)


def decide_quantity(
    *,
    mkr: str,
    code: str,
    name: str,
    raw_order_qty: Decimal,
    first_shipped_qty: Decimal,
    sales_note_order_qty: Decimal | None,
    backorder_order_qty: Decimal = ZERO,
    backorder_shipped_qty: Decimal = ZERO,
    shipped_fallback: bool = False,
    backorders_separated: bool = False,
) -> QuantityDecision:
    """Apply the six business rules in their required precedence order."""
    context = RuleContext(
        mkr=mkr,
        code=code,
        raw_order_qty=raw_order_qty,
        first_shipped_qty=first_shipped_qty,
        sales_note_order_qty=sales_note_order_qty,
        backorder_order_qty=backorder_order_qty,
        backorder_shipped_qty=backorder_shipped_qty,
        attempted_rule="입력 검증",
    )
    for value, label in (
        (raw_order_qty, "최초 Order QTY"),
        (first_shipped_qty, "최초 Shipped QTY"),
        (backorder_order_qty, "K열 합계"),
        (backorder_shipped_qty, "L열 합계"),
    ):
        _ensure_nonnegative(value, label, context)
    if sales_note_order_qty is not None:
        _ensure_nonnegative(sales_note_order_qty, "SALES NOTE Order QTY", context)
    if backorder_shipped_qty > backorder_order_qty:
        raise QuantityRuleError("L열 합계가 K열 합계보다 큽니다.", context)

    rule_id: str
    shipped_adjusted = False
    shipped_qty: Decimal | None = None

    if backorders_separated:
        if shipped_fallback:
            bad = RuleContext(
                **{**context.__dict__, "attempted_rule": "R10_CURRENT_ORDER_SALES_NOTE"}
            )
            raise QuantityRuleError(
                "현재 MKR 품목의 실제 Shipped QTY를 Sales Note에서 찾지 못했습니다.",
                bad,
            )
        rule_id = "R10_CURRENT_ORDER_SALES_NOTE"
        order_qty = raw_order_qty
        shipped_qty = first_shipped_qty
        if shipped_qty > order_qty:
            bad = RuleContext(**{**context.__dict__, "attempted_rule": rule_id})
            raise QuantityRuleError(
                "Sales Note의 Shipped QTY가 현재 MKR Order QTY보다 큽니다.", bad
            )
    elif (
        backorder_shipped_qty > ZERO
        and raw_order_qty == backorder_shipped_qty
        and first_shipped_qty == backorder_shipped_qty
        and backorder_order_qty >= backorder_shipped_qty
    ):
        rule_id = "R1_BACKORDER_ONLY"
        order_qty = ZERO
        shipped_qty = ZERO
    else:
        if backorder_order_qty == ZERO:
            rule_id = "R0_NO_BACKORDER"
            order_qty = raw_order_qty
        elif sales_note_order_qty is not None and raw_order_qty == sales_note_order_qty:
            rule_id = "R2_ALREADY_NET"
            order_qty = raw_order_qty
        elif sales_note_order_qty is not None and raw_order_qty == sales_note_order_qty + backorder_order_qty:
            rule_id = "R3_RAW_INCLUDES_K"
            order_qty = sales_note_order_qty
        elif sales_note_order_qty is not None and sales_note_order_qty == raw_order_qty + backorder_shipped_qty:
            rule_id = "R4_SALES_INCLUDES_L"
            order_qty = raw_order_qty
        elif sales_note_order_qty is not None and raw_order_qty == sales_note_order_qty + backorder_shipped_qty:
            rule_id = "R5_RAW_INCLUDES_L"
            order_qty = sales_note_order_qty
        elif sales_note_order_qty is not None and sales_note_order_qty == raw_order_qty + backorder_order_qty:
            rule_id = "R6_SALES_INCLUDES_K"
            order_qty = raw_order_qty
        elif sales_note_order_qty is None:
            if raw_order_qty >= backorder_order_qty:
                rule_id = "R7_MISSING_SALES_ORDER"
                order_qty = raw_order_qty - backorder_order_qty
            elif (
                backorder_shipped_qty > ZERO
                and raw_order_qty >= backorder_shipped_qty
                and (
                    first_shipped_qty == backorder_shipped_qty
                    or raw_order_qty == first_shipped_qty
                )
            ):
                candidate_order = raw_order_qty - backorder_shipped_qty
                candidate_shipped = first_shipped_qty - backorder_shipped_qty
                if ZERO <= candidate_shipped <= candidate_order:
                    rule_id = "R8_MISSING_SALES_INCLUDES_L"
                    order_qty = candidate_order
                    shipped_qty = candidate_shipped
                    shipped_adjusted = True
                else:
                    bad = RuleContext(**{**context.__dict__, "attempted_rule": "R8_MISSING_SALES_INCLUDES_L"})
                    raise QuantityRuleError("Sales Note 누락 O-L/D-L 교차검증에 실패했습니다.", bad)
            else:
                bad = RuleContext(**{**context.__dict__, "attempted_rule": "R7~R8"})
                raise QuantityRuleError("Sales Note 누락 O-K/O-L 계산 근거가 부족합니다.", bad)
        else:
            candidate_order = raw_order_qty - backorder_order_qty
            candidate_shipped = first_shipped_qty - backorder_shipped_qty
            anchored = (
                first_shipped_qty == backorder_shipped_qty
                or raw_order_qty == first_shipped_qty
            )
            if (
                backorder_shipped_qty > ZERO
                and anchored
                and candidate_order >= ZERO
                and ZERO <= candidate_shipped <= candidate_order
            ):
                rule_id = "R9_SALES_OUTLIER_KL_CROSSCHECK"
                order_qty = candidate_order
                shipped_qty = candidate_shipped
                shipped_adjusted = True
            else:
                bad = RuleContext(**{**context.__dict__, "attempted_rule": "R1~R9"})
                raise QuantityRuleError("어느 수량 판정 관계에도 일치하지 않습니다.", bad)

        if order_qty < ZERO:
            bad = RuleContext(**{**context.__dict__, "attempted_rule": rule_id})
            raise QuantityRuleError("계산된 신규 Order QTY가 음수입니다.", bad)

        if shipped_qty is None:
            shipped_qty = first_shipped_qty
        if shipped_fallback:
            shipped_adjusted = shipped_adjusted or shipped_qty != order_qty
            shipped_qty = order_qty
        elif shipped_qty > order_qty and backorder_shipped_qty > ZERO:
            adjusted = shipped_qty - backorder_shipped_qty
            if ZERO <= adjusted <= order_qty:
                shipped_qty = adjusted
                shipped_adjusted = True
        if shipped_qty > order_qty:
            bad = RuleContext(**{**context.__dict__, "attempted_rule": rule_id})
            raise QuantityRuleError("조정 후에도 Shipped QTY가 신규 Order QTY보다 큽니다.", bad)

    return QuantityDecision(
        mkr=mkr,
        code=code,
        name=name,
        raw_order_qty=raw_order_qty,
        first_shipped_qty=first_shipped_qty,
        sales_note_order_qty=sales_note_order_qty,
        backorder_order_qty=backorder_order_qty,
        backorder_shipped_qty=backorder_shipped_qty,
        order_qty=order_qty,
        shipped_qty=shipped_qty,
        rule_id=rule_id,
        shipped_fallback=shipped_fallback,
        shipped_adjusted=shipped_adjusted,
    )


def build_decisions(
    mkr: str,
    order_items: dict[str, OrderItem],
    sales_evidence: dict[str, SalesNoteEvidence],
    backorders: Iterable[BackorderSnapshot],
    *,
    sales_note_available: bool | None = None,
    order_source: str = "",
    sales_note_source: str = "",
) -> list[QuantityDecision]:
    if sales_note_available is None:
        sales_note_available = bool(sales_evidence)
    totals: dict[str, list[Decimal]] = defaultdict(lambda: [ZERO, ZERO])
    for item in backorders:
        totals[item.code][0] += item.order_qty
        totals[item.code][1] += item.shipped_qty

    decisions: list[QuantityDecision] = []
    for code in sorted(order_items):
        item = order_items[code]
        evidence = sales_evidence.get(code)
        k_qty, l_qty = totals[code]

        # A missing Sales Note is an unknown shipment, not a shipment equal to
        # the order.  Preserve the current order state and leave shipment and
        # shortage blank until evidence arrives.
        if not sales_note_available:
            if item.backorders_separated:
                order_qty = item.quantity
            elif item.quantity >= k_qty:
                order_qty = item.quantity - k_qty
            else:
                bad = RuleContext(
                    mkr=mkr,
                    code=code,
                    raw_order_qty=item.quantity,
                    first_shipped_qty=ZERO,
                    sales_note_order_qty=None,
                    backorder_order_qty=k_qty,
                    backorder_shipped_qty=l_qty,
                    attempted_rule="R11_PENDING_SALES_NOTE",
                )
                raise QuantityRuleError(
                    "Sales Note 미수신 상태에서 현재 주문수량과 백오더를 분리할 수 없습니다.",
                    bad,
                )
            decisions.append(
                QuantityDecision(
                    mkr=mkr,
                    code=code,
                    name=item.name,
                    raw_order_qty=item.quantity,
                    first_shipped_qty=None,
                    sales_note_order_qty=None,
                    backorder_order_qty=k_qty,
                    backorder_shipped_qty=l_qty,
                    order_qty=order_qty,
                    shipped_qty=None,
                    rule_id="R11_PENDING_SALES_NOTE",
                    origin=item.origin,
                    record_status="PENDING_SALES_NOTE",
                    order_source=item.source or order_source,
                    sales_note_source="",
                )
            )
            continue

        # A code omitted from a successfully parsed latest Sales Note has no
        # current shipment.  This is different from having no Sales Note at
        # all, and therefore resolves to an explicit zero.
        omitted_from_parsed_sales_note = evidence is None
        shipped_fallback = evidence is not None and evidence.shipped_qty is None
        first_shipped = ZERO if omitted_from_parsed_sales_note else (
            item.quantity if shipped_fallback else evidence.shipped_qty
        )
        if (
            item.backorders_separated
            and not shipped_fallback
            and evidence is not None
            and evidence.shipped_qty is not None
        ):
            # The Sales Note lists both current-MKR shipments and historical
            # backorder shipments.  Allocate the independently identified
            # backorder shipped total first; only the remainder belongs to
            # the current order.  This also handles a current item omitted
            # from the Sales Note while the same code appears as backorder.
            if evidence.shipped_qty < l_qty:
                bad = RuleContext(
                    mkr=mkr,
                    code=code,
                    raw_order_qty=item.quantity,
                    first_shipped_qty=evidence.shipped_qty,
                    sales_note_order_qty=evidence.order_qty,
                    backorder_order_qty=k_qty,
                    backorder_shipped_qty=l_qty,
                    attempted_rule="SALES_TOTAL_MINUS_BACKORDER_L",
                )
                raise QuantityRuleError(
                    "Sales Note 출하 합계가 확인된 백오더 출하 합계보다 작습니다.",
                    bad,
                )
            first_shipped = evidence.shipped_qty - l_qty
        sales_order = None if evidence is None else evidence.order_qty
        decision = decide_quantity(
            mkr=mkr,
            code=code,
            name=item.name,
            raw_order_qty=item.quantity,
            first_shipped_qty=first_shipped,
            sales_note_order_qty=sales_order,
            backorder_order_qty=k_qty,
            backorder_shipped_qty=l_qty,
            shipped_fallback=shipped_fallback,
            backorders_separated=item.backorders_separated,
        )
        decisions.append(
            replace(
                decision,
                origin=item.origin,
                record_status=("OMITTED_FROM_SALES_NOTE" if omitted_from_parsed_sales_note else "CURRENT"),
                order_source=item.source or order_source,
                sales_note_source=sales_note_source,
            )
        )

    # The Sales Note is shipment evidence.  After independently identified
    # historical backorders are allocated first, a positive remainder is a
    # current-MKR item even when no order workbook row was found.
    for code in sorted(set(sales_evidence).difference(order_items)):
        evidence = sales_evidence[code]
        k_qty, l_qty = totals[code]
        if evidence.shipped_qty is None:
            continue
        if evidence.shipped_qty < l_qty:
            bad = RuleContext(
                mkr=mkr,
                code=code,
                raw_order_qty=ZERO,
                first_shipped_qty=evidence.shipped_qty,
                sales_note_order_qty=evidence.order_qty,
                backorder_order_qty=k_qty,
                backorder_shipped_qty=l_qty,
                attempted_rule="R12_SALES_NOTE_ONLY",
            )
            raise QuantityRuleError(
                "Sales Note 전용 품목의 출하량이 확인된 백오더 출하량보다 작습니다.",
                bad,
            )
        current_shipped = evidence.shipped_qty - l_qty
        if current_shipped == ZERO:
            continue
        current_order: Decimal | None = None
        if evidence.order_qty is not None and evidence.order_qty >= k_qty:
            candidate = evidence.order_qty - k_qty
            if candidate >= current_shipped:
                current_order = candidate
        decisions.append(
            QuantityDecision(
                mkr=mkr,
                code=code,
                name=evidence.name,
                raw_order_qty=None,
                first_shipped_qty=evidence.shipped_qty,
                sales_note_order_qty=evidence.order_qty,
                backorder_order_qty=k_qty,
                backorder_shipped_qty=l_qty,
                order_qty=current_order,
                shipped_qty=current_shipped,
                rule_id="R12_SALES_NOTE_ONLY",
                origin="Unknown",
                record_status="SALES_NOTE_ONLY",
                order_source="",
                sales_note_source=sales_note_source,
            )
        )
    decisions.sort(key=lambda item: item.code)
    return decisions
