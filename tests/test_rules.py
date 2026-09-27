from decimal import Decimal as D
import unittest

from mkr_sync.errors import QuantityRuleError
from mkr_sync.rules import decide_quantity


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


if __name__ == "__main__":
    unittest.main()
