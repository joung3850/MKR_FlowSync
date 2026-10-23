from decimal import Decimal as D
import unittest

from mkr_sync.workbook_index import (
    WorkbookRecord,
    build_workbook_index,
    mkr_from_sheet_name,
    normalize_code,
)


def current(mkr, code, order, shipped, shortage, *, name="Product", origin="Japan", row=10):
    return WorkbookRecord(
        current_mkr=mkr,
        code=code,
        name=name,
        origin=origin,
        record_type="CURRENT",
        order_qty=order,
        shipped_qty=shipped,
        shortage_qty=shortage,
        row_number=row,
    )


def backorder(
    mkr,
    source_mkr,
    code,
    order,
    shipped,
    shortage,
    *,
    name="Product",
    origin="Japan",
    row=10,
):
    return WorkbookRecord(
        current_mkr=mkr,
        backorder_no=source_mkr,
        code=code,
        name=name,
        origin=origin,
        record_type="BACKORDER",
        order_qty=order,
        shipped_qty=shipped,
        shortage_qty=shortage,
        row_number=row,
    )


class WorkbookIndexTests(unittest.TestCase):
    def test_mkr_lookup_combines_current_and_backorder_shipped(self):
        index = build_workbook_index(
            [
                current("MKR60/26", "760254", D("50"), D("40"), D("10")),
                backorder(
                    "MKR60/26", "MKR55/26", "760254", D("100"), D("20"), D("80")
                ),
            ]
        )
        row = index.lookup_rows[0]
        self.assertEqual(row.new_order_qty, D("50"))
        self.assertEqual(row.backorder_order_qty, D("100"))
        self.assertEqual(row.total_imported_shipped_qty, D("60"))
        self.assertEqual(row.data_status, "완료")

    def test_product_totals_sum_shipments_but_keep_latest_backorder_snapshot(self):
        index = build_workbook_index(
            [
                current("MKR55/26", "760254", D("100"), D("80"), D("20")),
                backorder(
                    "MKR60/26", "MKR55/26", "760254", D("100"), D("10"), D("10")
                ),
                backorder(
                    "MKR61/26", "MKR55/26", "760254", D("100"), D("10"), D("0")
                ),
            ]
        )
        row = index.product_rows[0]
        self.assertEqual(row.new_order_qty, D("100"))
        self.assertEqual(row.new_shipped_qty, D("80"))
        self.assertEqual(row.backorder_shipped_qty, D("20"))
        self.assertEqual(row.total_imported_shipped_qty, D("100"))
        self.assertEqual(row.latest_backorder_order_qty, D("100"))
        self.assertEqual(row.latest_backorder_shortage_qty, D("0"))
        self.assertEqual(row.mkr_count, 3)

    def test_duplicate_backorder_key_in_same_mkr_uses_last_row(self):
        index = build_workbook_index(
            [
                backorder(
                    "MKR60/26", "MKR55/26", "760254", D("100"), D("10"), D("90"), row=10
                ),
                backorder(
                    "MKR60/26", "MKR55/26", "760254", D("100"), D("20"), D("80"), row=11
                ),
            ]
        )
        self.assertEqual(index.lookup_rows[0].backorder_shipped_qty, D("20"))
        self.assertEqual(index.product_rows[0].backorder_shipped_qty, D("20"))

    def test_missing_shipped_is_not_converted_to_zero(self):
        index = build_workbook_index(
            [current("MKR60/26", "760254", D("50"), None, None)]
        )
        lookup = index.lookup_rows[0]
        summary = index.product_rows[0]
        self.assertIsNone(lookup.new_shipped_qty)
        self.assertIsNone(lookup.total_imported_shipped_qty)
        self.assertIn("신규 Shipped 1건 미확인", lookup.data_status)
        self.assertIsNone(summary.total_imported_shipped_qty)
        self.assertIn("신규 Shipped 1건 미확인", summary.data_status)

    def test_latest_name_and_combined_origin_are_preserved(self):
        index = build_workbook_index(
            [
                current(
                    "MKR60/26", "760254", D("10"), D("10"), D("0"),
                    name="Old Name", origin="Japan"
                ),
                current(
                    "MKR61/26", "760254", D("20"), D("20"), D("0"),
                    name="New Name", origin="Thailand"
                ),
            ]
        )
        row = index.product_rows[0]
        self.assertEqual(row.name, "New Name")
        self.assertEqual(row.origin, "Japan, Thailand")

    def test_code_and_sheet_name_normalization(self):
        self.assertEqual(normalize_code(1234.0), "001234")
        self.assertEqual(normalize_code("760254"), "760254")
        self.assertEqual(mkr_from_sheet_name("MKR65_26-2"), "MKR65/26-2")
        self.assertIsNone(mkr_from_sheet_name("품목합계"))


if __name__ == "__main__":
    unittest.main()
