from datetime import datetime
from decimal import Decimal as D
import unittest

from mkr_sync.errors import DataValidationError
from mkr_sync.parsers import (
    attachment_kind,
    current_message_body,
    find_current_target_mkrs,
    find_target_mkrs,
    merge_sales_evidence,
    parse_backorders,
    parse_order_matrix,
    parse_date_near_keywords,
    parse_sales_note,
    validate_mkr_match,
)


class ParserTests(unittest.TestCase):
    def test_backorder_control_cells_keep_latest_duplicate_snapshot(self):
        text = (
            "Back Order QTY\a"
            "MKR48/26\a760383\aProduct Name\a120\a100\a20\a"
            "MKR48/26\a760383\aProduct Name\a120\a110\a10"
        )
        parsed = parse_backorders(text, "MKR56/26", datetime(2026, 8, 20), "fixture")
        item = parsed[("MKR48/26", "760383")]
        self.assertEqual(item.order_qty, D("120"))
        self.assertEqual(item.shipped_qty, D("110"))
        self.assertEqual(item.shortage_qty, D("10"))

    def test_backorder_cell_parser_keeps_number_at_end_of_product_name(self):
        text = (
            "欠品時No.\aCode\aName\aOrder QTY\aBack Order QTY\a"
            "MKR21/26-1\a558358\aAujua DIORUM SHAMPOO,9mL x 5\a120\a120"
        )
        parsed = parse_backorders(text, "MKR50/26", datetime(2026, 9, 8), "fixture")
        item = parsed[("MKR21/26-1", "558358")]
        self.assertEqual(item.name, "Aujua DIORUM SHAMPOO,9mL x 5")
        self.assertEqual((item.order_qty, item.shipped_qty), (D("120"), D("120")))

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

    def test_other_subject_mkr_does_not_route_from_backorder_body(self):
        targets = {"MKR48/26", "MKR50/26"}
        self.assertEqual(
            find_current_target_mkrs(
                "RE: ORDER MKR56/26",
                "Back Order MKR48/26 760245 Product 720 720",
                targets,
            ),
            [],
        )
        self.assertEqual(
            find_current_target_mkrs("Shipment update", "Current MKR50/26", targets),
            ["MKR50/26"],
        )

    def test_revision_and_attachment_classification(self):
        self.assertEqual(attachment_kind("Rev_MKR56_26_Order.xlsx"), ("order", True))
        self.assertEqual(attachment_kind("MKR56_26_SALES NOTE.pdf"), ("sales_note", False))

    def test_logistics_spreadsheets_are_not_original_orders(self):
        excluded = [
            "Lot No. List MKR49_26.xlsx",
            "PALLET DETAIL MKR49_26.xlsx",
            "Container Weighting Calculation- MKR48_26-1.xlsx",
            "INVOICE MKR48_26.xls",
            "TSI_MKR48_26.xlsm",
            "MKR48_26-1 (TRADE SHIPPING)-rev.xls",
            "【出荷日変更】MKR49_26-1.xlsx",
        ]
        for filename in excluded:
            with self.subTest(filename=filename):
                kind, _revision = attachment_kind(filename)
                self.assertIsNone(kind)

    def test_order_matrix_fallback_rejects_container_identifiers(self):
        parsed, fallback = parse_order_matrix(
            [
                ["011171", "SKHU9959469", "108909"],
                ["108837", "40 FT'", "41.14"],
            ]
        )
        self.assertTrue(fallback)
        self.assertEqual(parsed, {})

    def test_mkr_mismatch_stops(self):
        with self.assertRaises(DataValidationError):
            validate_mkr_match("MKR56/26", ["MKR48/26"], "첨부파일")

    def test_target_mkr_with_historical_backorder_mkrs_is_allowed(self):
        validate_mkr_match(
            "MKR49/26-1",
            ["MKR16/26-1", "MKR20/26-2", "MKR49/26-1"],
            "주문서",
        )

    def test_english_month_first_date(self):
        value = parse_date_near_keywords(
            "ETA: September 14, 2026",
            r"\bETA\b",
            datetime(2026, 8, 20),
        )
        self.assertEqual(value.isoformat(), "2026-09-14")


if __name__ == "__main__":
    unittest.main()
