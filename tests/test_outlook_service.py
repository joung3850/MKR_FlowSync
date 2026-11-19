import unittest

from mkr_sync.outlook_service import _message_target_mkrs


class OutlookServiceTests(unittest.TestCase):
    def test_attachment_filename_can_identify_target_mkr(self):
        targets = {"MKR62/26", "MKR63/26-1"}

        result = _message_target_mkrs(
            "Container documents",
            "Please see attached files.",
            [["MKR62/26", "MKR63/26-1"]],
            targets,
        )

        self.assertEqual(result, ["MKR62/26", "MKR63/26-1"])


if __name__ == "__main__":
    unittest.main()
