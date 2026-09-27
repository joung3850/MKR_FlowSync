import unittest

from mkr_sync.mkr_parser import extract_mkr_numbers, to_excel_sheet_name


class MkrParserTests(unittest.TestCase):
    def test_extracts_all_eight_numbers_from_production_description(self):
        text = """MKR56/26・・・・・タイ産・アンソン向け・普通品
MKR56/26-1・・・タイ産・アンソン向け・専用品
MKR56/26-2・・・タイ産・アンソン向け・専用品（P/O No.：MKR-20260910TG-A(T）
MKR57/26・・・・・タイ産・ピョンテク向け・普通品
MKR58/26・・・・・日本産・アンソン向け・普通品
MKR58/26-1・・・日本産・アンソン向け・危険品
MKR59/26・・・・・日本産・ピョンテク向け・普通品
MKR59/26-1・・・日本産・ピョンテク向け・危険品"""
        self.assertEqual(
            extract_mkr_numbers(text),
            [
                "MKR56/26", "MKR56/26-1", "MKR56/26-2", "MKR57/26",
                "MKR58/26", "MKR58/26-1", "MKR59/26", "MKR59/26-1",
            ],
        )

    def test_extracts_identifiers_from_japanese_description(self):
        text = """MKR56/26・・・・・タイ産・アンソン向け・普通品
MKR56/26-1・・・タイ産・アンソン向け・専用品
MKR56/26-2・・・専用品（P/O No.：MKR-20260910TG-A(T）
MKR57/26・・・・・タイ産・ピョンテク向け・普通品"""
        self.assertEqual(
            extract_mkr_numbers(text),
            ["MKR56/26", "MKR56/26-1", "MKR56/26-2", "MKR57/26"],
        )

    def test_deduplicates_and_normalizes_case(self):
        self.assertEqual(extract_mkr_numbers("mkr56/26 MKR56/26 MKR 57 / 26"), ["MKR56/26", "MKR57/26"])

    def test_does_not_match_po_identifier(self):
        self.assertEqual(extract_mkr_numbers("P/O No.: MKR-20260910TG-A(T)"), [])

    def test_converts_slash_for_excel_sheet_name(self):
        self.assertEqual(to_excel_sheet_name("MKR56/26"), "MKR56_26")
        self.assertEqual(to_excel_sheet_name("MKR56/26-1"), "MKR56_26-1")


if __name__ == "__main__":
    unittest.main()
