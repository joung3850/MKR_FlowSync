import logging
import os
from pathlib import Path
import shutil
import tempfile
import unittest

from mkr_sync.office import sha256_file
from mkr_sync.sheet_creator import create_mkr_sheets


@unittest.skipUnless(os.environ.get("MKR_RUN_COM_TESTS") == "1", "Office COM integration test")
class ExternalTemplateSheetTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        project = Path(__file__).resolve().parents[1]
        self.source = project / "MKR_TEST.xlsx"
        self.template = project / "MKR_TEMPLATE.xlsx"
        self.target = self.root / "target.xlsx"
        shutil.copy2(self.source, self.target)
        self.logger = logging.getLogger("sheet-creator-integration")

    def tearDown(self):
        self.temp.cleanup()

    def test_external_template_creates_multiple_sheets_without_modifying_template(self):
        template_hash = sha256_file(self.template)
        result = create_mkr_sheets(
            self.target,
            ["MKR57/26", "MKR58/26-1"],
            template_workbook=self.template,
            backup_dir=self.root / "Backup",
            logger=self.logger,
        )
        self.assertEqual(result["created"], ["MKR57_26", "MKR58_26-1"])
        self.assertEqual(sha256_file(self.template), template_hash)

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(self.target), 0, True)
            names = {
                str(workbook.Worksheets.Item(index).Name)
                for index in range(1, int(workbook.Worksheets.Count) + 1)
            }
            self.assertIn("MKR57_26", names)
            self.assertIn("MKR58_26-1", names)
            self.assertNotIn("MKR_TEMPLATE", names)
            self.assertEqual(workbook.Worksheets.Item("MKR57_26").Range("D3").Value2, "MKR57/26")
        finally:
            if workbook is not None:
                workbook.Close(False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    unittest.main()
