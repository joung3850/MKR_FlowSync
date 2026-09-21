from datetime import datetime
from decimal import Decimal as D
import unittest

from mkr_sync.errors import DataValidationError
from mkr_sync.parsers import (
    attachment_kind,
    current_message_body,
    find_target_mkrs,
    merge_sales_evidence,
    parse_backorders,
    parse_order_matrix,
    parse_date_near_keywords,
    parse_sales_note,
    validate_mkr_match,
)


class ParserTests(unittest.TestCase):
    def test_backorder_control_cells_and_duplicate_rows_are_summed(self):
        text = (
            "Back Order QTY\a"
            "MKR48/26\a760383\aProduct Name\a120\a100\a20\a"
            "MKR48/26\a760383\aProduct Name\a120\a140\a0"
        )
        parsed = parse_backorders(text, "MKR56/26", datetime(2026, 8, 20), "fixture")
        item = parsed[("MKR48/26", "760383")]
        self.assertEqual(item.order_qty, D("240"))
        self.assertEqual(item.shipped_qty, D("240"))
        self.assertEqual(item.shortage_qty, D("0"))

    def test_sales_note_full_row_and_cell_row(self):
        full = parse_sales_note("760526 Product Name 600 120")
        self.assertEqual(full["760526"].order_qty, D("600"))
        self.assertEqual(full["760526"].shipped_qty, D("120"))
        cells = parse_sales_note("Code\a760383\aActual Product\a2040")
        self.assertIsNone(cells["760383"].order_qty)
        self.assertEqual(cells["760383"].shipped_qty, D("2040"))

    def test_pdf_evidence_wins_and_body_fills_missing_value(self):
        pdf = parse_sales_note("760383 Product Name 1800")
        body = parse_sales_note("760383 Product Name 1800 1750")
        merged = merge_sales_evidence(pdf, body)
        self.assertEqual(merged["760383"].order_qty, D("1800"))
        self.assertEqual(merged["760383"].shipped_qty, D("1800"))

    def test_order_matrix_header_parser_sums_duplicate_codes(self):
        matrix = [
            ["Code", "Name", "Order QTY"],
            ["417133", "Product", "28000"],
            ["417133", "Product", "800"],
        ]
        parsed, fallback = parse_order_matrix(matrix)
        self.assertFalse(fallback)
        self.assertEqual(parsed["417133"].quantity, D("28800"))

    def test_order_matrix_fallback(self):
        parsed, fallback = parse_order_matrix([["760526", "Product", "180"]])
        self.assertTrue(fallback)
        self.assertEqual(parsed["760526"].quantity, D("180"))

    def test_subject_is_only_current_mkr_source(self):
        targets = {"MKR48/26", "MKR56/26"}
        self.assertEqual(find_target_mkrs("Sales Note MKR56_26", targets), ["MKR56/26"])
        body = current_message_body(
            "Current text\nBack Order MKR48/26\n-----Original Message-----\nMKR56/26 old quote"
        )
        self.assertNotIn("old quote", body)

    def test_revision_and_attachment_classification(self):
        self.assertEqual(attachment_kind("Rev_MKR56_26_Order.xlsx"), ("order", True))
        self.assertEqual(attachment_kind("MKR56_26_SALES NOTE.pdf"), ("sales_note", False))

    def test_mkr_mismatch_stops(self):
        with self.assertRaises(DataValidationError):
            validate_mkr_match("MKR56/26", ["MKR48/26"], "첨부파일")

    def test_english_month_first_date(self):
        value = parse_date_near_keywords(
            "ETA: September 14, 2026",
            r"\bETA\b",
            datetime(2026, 8, 20),
        )
        self.assertEqual(value.isoformat(), "2026-09-14")


if __name__ == "__main__":
    unittest.main()
