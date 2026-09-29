from datetime import datetime
from decimal import Decimal as D
import unittest

from mkr_sync.errors import QuantityRuleError
from mkr_sync.models import BackorderSnapshot, OrderItem, SalesNoteEvidence
from mkr_sync.rules import build_decisions, decide_quantity


class QuantityRuleTests(unittest.TestCase):
    def decide(self, **overrides):
        values = {
            "mkr": "MKR56/26",
            "code": "999999",
            "name": "Test Item",
            "raw_order_qty": D("100"),
            "first_shipped_qty": D("80"),
            "sales_note_order_qty": D("100"),
            "backorder_order_qty": D("0"),
            "backorder_shipped_qty": D("0"),
        }
        values.update(overrides)
        return decide_quantity(**values)

    def test_417133_rule3(self):
        result = self.decide(
            code="417133",
            raw_order_qty=D("87336"),
            first_shipped_qty=D("28800"),
            sales_note_order_qty=D("28800"),
            backorder_order_qty=D("58536"),
        )
        self.assertEqual(result.rule_id, "R3_RAW_INCLUDES_K")
        self.assertEqual(result.order_qty, D("28800"))

    def test_760351_backorder_only_rule1_has_first_precedence(self):
        result = self.decide(
            code="760351",
            raw_order_qty=D("60"),
            first_shipped_qty=D("60"),
            sales_note_order_qty=D("60"),
            backorder_order_qty=D("1320"),
            backorder_shipped_qty=D("60"),
        )
        self.assertEqual(result.rule_id, "R1_BACKORDER_ONLY")
        self.assertEqual((result.order_qty, result.shipped_qty, result.shortage_qty), (D("0"), D("0"), D("0")))

    def test_760526_rule4(self):
        result = self.decide(
            code="760526",
            raw_order_qty=D("180"),
            first_shipped_qty=D("120"),
            sales_note_order_qty=D("600"),
            backorder_order_qty=D("480"),
            backorder_shipped_qty=D("420"),
        )
        self.assertEqual(result.rule_id, "R4_SALES_INCLUDES_L")
        self.assertEqual((result.order_qty, result.shipped_qty, result.shortage_qty), (D("180"), D("120"), D("60")))

    def test_760383_rule6_adjusts_shipped(self):
        result = self.decide(
            code="760383",
            raw_order_qty=D("2040"),
            first_shipped_qty=D("2040"),
            sales_note_order_qty=None,
            backorder_order_qty=D("240"),
            backorder_shipped_qty=D("240"),
        )
        self.assertEqual(result.rule_id, "R7_MISSING_SALES_ORDER")
        self.assertTrue(result.shipped_adjusted)
        self.assertEqual((result.order_qty, result.shipped_qty, result.shortage_qty), (D("1800"), D("1800"), D("0")))

    def test_rule2_already_net(self):
        result = self.decide(backorder_order_qty=D("40"), backorder_shipped_qty=D("20"))
        self.assertEqual(result.rule_id, "R2_ALREADY_NET")
        self.assertEqual(result.order_qty, D("100"))

    def test_rule5_raw_includes_l(self):
        result = self.decide(
            raw_order_qty=D("180"),
            first_shipped_qty=D("100"),
            sales_note_order_qty=D("120"),
            backorder_order_qty=D("100"),
            backorder_shipped_qty=D("60"),
        )
        self.assertEqual(result.rule_id, "R5_RAW_INCLUDES_L")
        self.assertEqual(result.order_qty, D("120"))

    def test_explicit_zero_sales_order_is_not_missing(self):
        result = self.decide(
            raw_order_qty=D("100"),
            first_shipped_qty=D("0"),
            sales_note_order_qty=D("0"),
            backorder_order_qty=D("100"),
        )
        self.assertEqual(result.rule_id, "R3_RAW_INCLUDES_K")
        self.assertEqual(result.order_qty, D("0"))

    def test_sales_mismatch_without_backorder_keeps_original_order(self):
        result = self.decide(sales_note_order_qty=D("70"))
        self.assertEqual(result.rule_id, "R0_NO_BACKORDER")
        self.assertEqual(result.order_qty, D("100"))

    def test_k_greater_than_raw_stops_when_not_rule1(self):
        with self.assertRaises(QuantityRuleError):
            self.decide(
                sales_note_order_qty=D("70"),
                backorder_order_qty=D("101"),
                backorder_shipped_qty=D("20"),
            )

    def test_shipped_above_order_stops(self):
        with self.assertRaises(QuantityRuleError):
            self.decide(first_shipped_qty=D("101"))

    def test_rule6_invalid_shipped_adjustment_stops(self):
        with self.assertRaises(QuantityRuleError):
            self.decide(
                raw_order_qty=D("100"),
                first_shipped_qty=D("150"),
                sales_note_order_qty=None,
                backorder_order_qty=D("20"),
                backorder_shipped_qty=D("10"),
            )

    def test_missing_sales_uses_final_order_as_shipped(self):
        result = self.decide(
            code="330883",
            raw_order_qty=D("960"),
            first_shipped_qty=D("960"),
            sales_note_order_qty=None,
            backorder_order_qty=D("480"),
            backorder_shipped_qty=D("180"),
            shipped_fallback=True,
        )
        self.assertEqual(result.rule_id, "R7_MISSING_SALES_ORDER")
        self.assertEqual((result.order_qty, result.shipped_qty), (D("480"), D("480")))

    def test_missing_sales_with_k_above_order_uses_l_crosscheck(self):
        result = self.decide(
            code="754303",
            raw_order_qty=D("1500"),
            first_shipped_qty=D("1500"),
            sales_note_order_qty=None,
            backorder_order_qty=D("2400"),
            backorder_shipped_qty=D("1200"),
            shipped_fallback=True,
        )
        self.assertEqual(result.rule_id, "R8_MISSING_SALES_INCLUDES_L")
        self.assertEqual((result.order_qty, result.shipped_qty), (D("300"), D("300")))

    def test_separated_current_order_uses_actual_sales_note_without_backorder_deduction(self):
        decisions = build_decisions(
            "MKR49/26",
            {"331383": OrderItem("331383", "ORDEVE 13-srMV", D("1800"), True)},
            {
                "331383": SalesNoteEvidence(
                    "331383", "ORDEVE 13-srMV", None, D("2160")
                )
            },
            [
                BackorderSnapshot(
                    "MKR49/26",
                    "MKR20/26-1",
                    "331383",
                    "ORDEVE 13-srMV",
                    D("1200"),
                    D("1080"),
                    datetime(2026, 9, 1),
                    "fixture",
                )
            ],
        )
        result = decisions[0]
        self.assertEqual(result.rule_id, "R10_CURRENT_ORDER_SALES_NOTE")
        self.assertEqual(
            (result.order_qty, result.shipped_qty, result.shortage_qty),
            (D("1800"), D("1080"), D("720")),
        )

    def test_mkr49_1_duplicate_code_sales_are_allocated_to_backorders_first(self):
        decisions = build_decisions(
            "MKR49/26-1",
            {"417133": OrderItem("417133", "NIGELLE HOLDFIT VEIL", D("28800"), True)},
            {
                "417133": SalesNoteEvidence(
                    "417133", "NIGELLE HOLDFIT VEIL", D("28800"), D("50040")
                )
            },
            [
                BackorderSnapshot(
                    "MKR49/26-1", "MKR20/26-2", "417133", "NIGELLE HOLDFIT VEIL",
                    D("29736"), D("29736"), datetime(2026, 9, 8), "fixture",
                ),
                BackorderSnapshot(
                    "MKR49/26-1", "MKR39/26-1", "417133", "NIGELLE HOLDFIT VEIL",
                    D("28800"), D("20304"), datetime(2026, 9, 8), "fixture",
                ),
            ],
        )
        result = decisions[0]
        self.assertEqual(
            (result.order_qty, result.shipped_qty, result.shortage_qty),
            (D("28800"), D("0"), D("28800")),
        )

    def test_nonseparated_duplicate_sales_total_can_resolve_backorder_only(self):
        result = self.decide(
            code="760475",
            raw_order_qty=D("1020"),
            first_shipped_qty=D("1020"),
            sales_note_order_qty=None,
            backorder_order_qty=D("1020"),
            backorder_shipped_qty=D("1020"),
        )
        self.assertEqual(result.rule_id, "R1_BACKORDER_ONLY")
        self.assertEqual((result.order_qty, result.shipped_qty), (D("0"), D("0")))

    def test_separated_current_order_requires_sales_note_shipped_value(self):
        with self.assertRaises(QuantityRuleError):
            self.decide(
                sales_note_order_qty=None,
                shipped_fallback=True,
                backorders_separated=True,
            )

    def test_item_omitted_from_valid_sales_note_is_zero_shipped(self):
        decisions = build_decisions(
            "MKR49/26",
            {
                "302509": OrderItem("302509", "Missing item", D("3600"), True),
                "331383": OrderItem("331383", "Present item", D("1800"), True),
            },
            {
                "331383": SalesNoteEvidence(
                    "331383", "Present item", None, D("1080")
                )
            },
            [],
        )
        by_code = {item.code: item for item in decisions}
        self.assertEqual(by_code["302509"].shipped_qty, D("0"))
        self.assertFalse(by_code["302509"].shipped_fallback)

    def test_missing_sales_note_keeps_shipment_and_shortage_unknown(self):
        decisions = build_decisions(
            "MKR55/26",
            {"162365": OrderItem("162365", "Added item", D("72"), origin="Japan")},
            {},
            [],
            sales_note_available=False,
        )
        result = decisions[0]
        self.assertEqual(result.order_qty, D("72"))
        self.assertIsNone(result.shipped_qty)
        self.assertIsNone(result.shortage_qty)
        self.assertEqual(result.record_status, "PENDING_SALES_NOTE")

    def test_sales_note_only_item_is_kept_after_backorder_allocation(self):
        decisions = build_decisions(
            "MKR55/26",
            {},
            {"162365": SalesNoteEvidence("162365", "Added item", D("72"), D("72"))},
            [],
            sales_note_available=True,
            sales_note_source="latest.pdf",
        )
        result = decisions[0]
        self.assertEqual(result.rule_id, "R12_SALES_NOTE_ONLY")
        self.assertEqual((result.order_qty, result.shipped_qty), (D("72"), D("72")))
        self.assertEqual(result.record_status, "SALES_NOTE_ONLY")


if __name__ == "__main__":
    unittest.main()
