from datetime import date, datetime
from decimal import Decimal as D
import logging
import os
from pathlib import Path
import shutil
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from mkr_sync.apply_service import apply_preview_results, deserialize_preview_results
from mkr_sync.errors import DataValidationError
from mkr_sync.errors import WorkbookCommitError
from mkr_sync.office import _validate_preserved_number_formats, sha256_file
from mkr_sync.web_api import WebApi


def preview_payload(target: Path) -> dict:
    return {
        "kind": "outlook_preview",
        "target_workbook": str(target.resolve()),
        "workbook_hash": sha256_file(target),
        "mkr_numbers": ["MKR56/26"],
        "senders": ["sender@example.com"],
        "start_date": "2026-08-01",
        "max_messages": 500,
        "errors": [],
        "warnings": [],
        "results": [
            {
                "mkr": "MKR56/26",
                "sheet": "MKR56_26",
                "order_date": "2026-08-01",
                "ship_date": "2026-08-03",
                "eta": "2026-08-08",
                "etd": "2026-08-04",
                "decisions": [
                    {
                        "code": "760526",
                        "name": "Test Product",
                        "r": "180",
                        "s": "120",
                        "p": "600",
                        "k": "480",
                        "l": "420",
                        "order_qty": "180",
                        "shipped_qty": "120",
                        "shortage_qty": "60",
                        "rule": "R4_SALES_INCLUDES_L",
                        "shipped_fallback": False,
                        "shipped_adjusted": False,
                    }
                ],
                "backorders": [
                    {
                        "source_mkr": "MKR48/26",
                        "code": "760526",
                        "name": "Test Product",
                        "order_qty": "480",
                        "shipped_qty": "420",
                        "shortage_qty": "60",
                        "received_at": "2026-08-20T10:30:00",
                        "source": "fixture",
                    }
                ],
            }
        ],
    }


class ApplyServiceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.target = self.root / "target.xlsx"
        self.target.write_bytes(b"test workbook placeholder")

    def tearDown(self):
        self.temp.cleanup()

    def test_deserialize_preview_results_preserves_excel_values(self):
        results = deserialize_preview_results(preview_payload(self.target))
        self.assertEqual(len(results), 1)
        result = results[0]
        self.assertEqual(result.order_date, date(2026, 8, 1))
        self.assertEqual(result.decisions[0].order_qty, D("180"))
        self.assertEqual(result.decisions[0].shortage_qty, D("60"))
        self.assertEqual(result.backorders[0].received_at, datetime(2026, 8, 20, 10, 30))

    def test_deserialize_pending_sales_note_keeps_unknown_quantities_blank(self):
        payload = preview_payload(self.target)
        item = payload["results"][0]["decisions"][0]
        item.update({"s": None, "p": None, "shipped_qty": None, "shortage_qty": None})
        results = deserialize_preview_results(payload)
        self.assertIsNone(results[0].decisions[0].shipped_qty)
        self.assertIsNone(results[0].decisions[0].shortage_qty)

    def test_changed_workbook_is_rejected_before_excel_automation(self):
        payload = preview_payload(self.target)
        self.target.write_bytes(b"changed after preview")
        with self.assertRaisesRegex(DataValidationError, "미리보기 이후"):
            apply_preview_results(
                self.root,
                self.target,
                payload,
                logging.getLogger("apply-test"),
                threading.Event(),
                lambda *_: None,
            )

    def test_web_api_applies_completed_preview_once(self):
        api = WebApi(self.root, logging.getLogger("api-apply-test"))
        api.target_workbook = self.target
        payload = preview_payload(self.target)
        preview_job_id = api.jobs.start("outlook", lambda *_: payload)
        for _ in range(100):
            if api.jobs.status(preview_job_id)["state"] == "completed":
                break
            time.sleep(0.01)

        applied = {
            "kind": "excel_apply",
            "updated_sheets": ["MKR56_26"],
            "warnings": [],
            "errors": [],
        }
        with patch("mkr_sync.web_api.apply_preview_results", return_value=applied):
            response = api.apply_preview({"preview_job_id": preview_job_id})
            self.assertTrue(response["ok"])
            for _ in range(100):
                if api.jobs.status(response["job_id"])["state"] == "completed":
                    break
                time.sleep(0.01)

        duplicate = api.apply_preview({"preview_job_id": preview_job_id})
        self.assertFalse(duplicate["ok"])
        self.assertIn("이미 Excel에 적용", duplicate["error"])

    def test_web_api_allows_successful_results_when_another_mkr_failed(self):
        api = WebApi(self.root, logging.getLogger("api-partial-apply-test"))
        api.target_workbook = self.target
        payload = preview_payload(self.target)
        payload["errors"] = [
            {"level": "error", "code": "DataValidationError", "message": "ambiguous", "mkr": "MKR55/26"}
        ]
        preview_job_id = api.jobs.start("outlook", lambda *_: payload)
        for _ in range(100):
            if api.jobs.status(preview_job_id)["state"] == "completed":
                break
            time.sleep(0.01)
        with patch("mkr_sync.web_api.apply_preview_results", return_value={"kind": "excel_apply"}):
            response = api.apply_preview({"preview_job_id": preview_job_id})
            for _ in range(100):
                if api.jobs.status(response["job_id"])["state"] == "completed":
                    break
                time.sleep(0.01)
        self.assertTrue(response["ok"])

    def test_html_exposes_excel_apply_control(self):
        html = (Path(__file__).resolve().parents[1] / "mkr_sync" / "web" / "index.html").read_text(
            encoding="utf-8"
        )
        self.assertIn('id="applyButton"', html)
        self.assertIn('api("apply_preview"', html)

    def test_html_defaults_outlook_search_to_ninety_days(self):
        html = (Path(__file__).resolve().parents[1] / "mkr_sync" / "web" / "index.html").read_text(
            encoding="utf-8"
        )
        self.assertIn("d.setDate(d.getDate()-90)", html)

    def test_template_integer_format_zero_is_accepted_when_preserved(self):
        formats = {"C10": "0", "D10": "0", "E10": "0"}
        _validate_preserved_number_formats("MKR52_26", formats, dict(formats))

    def test_changed_number_format_is_rejected_against_source_format(self):
        with self.assertRaisesRegex(WorkbookCommitError, "C10.*원본=0.*임시=#,##0"):
            _validate_preserved_number_formats(
                "MKR52_26",
                {"C10": "0"},
                {"C10": "#,##0"},
            )


@unittest.skipUnless(os.environ.get("MKR_RUN_COM_TESTS") == "1", "Office COM integration test")
class ApplyServiceComTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        source = Path(__file__).resolve().parents[1] / "MKR_TEST.xlsx"
        self.target = self.root / "target.xlsx"
        shutil.copy2(source, self.target)

    def tearDown(self):
        self.temp.cleanup()

    def test_serialized_preview_is_committed_to_expected_cells(self):
        payload = preview_payload(self.target)
        payload["mkr_numbers"] = ["MKR52/26"]
        payload["results"][0]["mkr"] = "MKR52/26"
        payload["results"][0]["sheet"] = "MKR52_26"
        payload["results"][0]["decisions"][0]["mkr"] = "MKR52/26"
        result = apply_preview_results(
            self.root,
            self.target,
            payload,
            logging.getLogger("apply-com-test"),
            threading.Event(),
            lambda *_: None,
        )
        self.assertEqual(result["updated_sheets"], ["MKR52_26"])
        self.assertTrue(Path(result["backup_path"]).is_file())
        self.assertTrue(Path(result["report_path"]).is_file())

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(self.target), 0, True)
            sheet = workbook.Worksheets.Item("MKR52_26")
            self.assertEqual(sheet.Range("D3").Value2, "MKR52/26")
            self.assertEqual(tuple(sheet.Range("A10:E10").Value2[0]), ("760526", "Test Product", 180.0, 120.0, 60.0))
            self.assertEqual(tuple(sheet.Range("H10:M10").Value2[0]), ("MKR48/26", "760526", "Test Product", 480.0, 420.0, 60.0))
            self.assertEqual(tuple(sheet.Range("K4:K7").Value2), (("2026-08-01",), ("2026-08-03",), ("2026-08-08",), ("2026-08-04",)))
        finally:
            if workbook is not None:
                workbook.Close(False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    unittest.main()
