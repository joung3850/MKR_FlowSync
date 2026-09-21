from datetime import date
from decimal import Decimal as D
import logging
import os
from pathlib import Path
import shutil
import tempfile
import unittest

from mkr_sync.config import AppConfig, REQUIRED_TARGETS
from mkr_sync.errors import WorkbookCommitError
from mkr_sync.models import BackorderSnapshot, MkrResult, QuantityDecision
from mkr_sync.office import WorkbookTransaction, assert_workbook_available, sha256_file


@unittest.skipUnless(os.environ.get("MKR_RUN_COM_TESTS") == "1", "Office COM integration test")
class ExcelTransactionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = Path(__file__).resolve().parents[1] / "MKR_TEST.xlsx"
        self.workbook = self.root / "MKR_TEST.xlsx"
        shutil.copy2(self.source, self.workbook)
        self.config = AppConfig(
            root=self.root,
            workbook=self.workbook,
            sender="hnagasawa@milbon.com",
            start_date=date(2026, 8, 1),
            max_messages=500,
            data_start_row=10,
            max_rows=500,
            targets=dict(REQUIRED_TARGETS),
        )
        self.logger = logging.getLogger("integration-test")
        decision = QuantityDecision(
            mkr="MKR56/26",
            code="760526",
            name="Integration Test",
            raw_order_qty=D("180"),
            first_shipped_qty=D("120"),
            sales_note_order_qty=D("600"),
            backorder_order_qty=D("480"),
            backorder_shipped_qty=D("420"),
            order_qty=D("180"),
            shipped_qty=D("120"),
            rule_id="R4_SALES_INCLUDES_L",
        )
        backorder = BackorderSnapshot(
            current_mkr="MKR56/26",
            source_mkr="MKR48/26",
            code="760526",
            name="Integration Test",
            order_qty=D("480"),
            shipped_qty=D("420"),
            received_at=__import__("datetime").datetime(2026, 8, 20),
            source="fixture",
        )
        self.results = [
            MkrResult(
                mkr="MKR56/26",
                sheet="MKR56_26",
                decisions=[decision],
                backorders=[backorder],
                order_date=date(2026, 8, 1),
            )
        ]

    def tearDown(self):
        self.temp.cleanup()

    def test_savecopyas_validate_and_commit_preserve_source_template(self):
        source_hash = sha256_file(self.source)
        original_copy_hash = sha256_file(self.workbook)
        transaction = WorkbookTransaction(self.config, self.logger)
        backup = transaction.prepare()
        self.assertEqual(sha256_file(backup), original_copy_hash)
        transaction.write_pending(self.results)
        transaction.validate_pending(self.results)
        transaction.commit()
        self.assertNotEqual(sha256_file(self.workbook), original_copy_hash)
        self.assertEqual(sha256_file(self.source), source_hash)

    def test_validation_failure_does_not_replace_original(self):
        original_hash = sha256_file(self.workbook)
        transaction = WorkbookTransaction(self.config, self.logger)
        transaction.prepare()
        transaction.write_pending(self.results)
        transaction.pending_path.unlink()
        with self.assertRaises(WorkbookCommitError):
            transaction.validate_pending(self.results)
        self.assertEqual(sha256_file(self.workbook), original_hash)

    def test_open_workbook_is_rejected(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(self.workbook), 0, False)
            with self.assertRaises(WorkbookCommitError):
                assert_workbook_available(self.workbook)
        finally:
            if workbook is not None:
                workbook.Close(False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    unittest.main()
