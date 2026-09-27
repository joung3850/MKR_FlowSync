from datetime import datetime, timedelta
from decimal import Decimal as D
from pathlib import Path
import unittest

from mkr_sync.errors import DataValidationError
from mkr_sync.models import AttachmentEvent, BackorderSnapshot
from mkr_sync.workflow import (
    merge_latest_backorders,
    original_order_candidates,
    select_first_successful_order,
    validate_capacity,
)


def event(entry_id, name, digest, minutes=0, revision=False):
    return AttachmentEvent(
        entry_id=entry_id,
        target_mkr="MKR56/26",
        received_at=datetime(2026, 8, 1) + timedelta(minutes=minutes),
        subject="MKR56/26",
        body="",
        kind="order",
        is_revision=revision,
        original_name=name,
        path=Path(name),
        sha256=digest,
    )


class WorkflowTests(unittest.TestCase):
    def test_order_candidates_include_later_files_and_deduplicate_hash(self):
        events = [
            event("first", "A.xlsx", "aaa"),
            event("first", "A-copy.xlsx", "aaa"),
            event("first", "B.xlsx", "bbb"),
            event("later", "later.xlsx", "ccc", minutes=10),
            event("first", "Rev_A.xlsx", "ddd", revision=True),
        ]
        selected = original_order_candidates(events)
        self.assertEqual({item.sha256 for item in selected}, {"aaa", "bbb", "ccc"})
        self.assertEqual(len(selected), 3)

    def test_first_successful_order_skips_empty_and_uses_most_complete_at_same_time(self):
        empty = event("first", "Lot No.xlsx", "aaa")
        small = event("second", "B.xlsx", "bbb", minutes=10)
        complete = event("second", "A.xlsx", "ccc", minutes=10)
        later = event("later", "later.xlsx", "ddd", minutes=20)
        selected = select_first_successful_order(
            [
                (small, {"1": object()}),
                (complete, {"1": object(), "2": object()}),
                (later, {"1": object(), "2": object(), "3": object()}),
            ]
        )
        self.assertEqual(selected[0], complete)
        self.assertNotEqual(selected[0], empty)

    def test_latest_backorder_snapshot_replaces_old_and_excludes_current_mkr(self):
        old = BackorderSnapshot("MKR56/26", "MKR48/26", "760010", "P", D("100"), D("10"), datetime(2026, 8, 1), "old")
        new = BackorderSnapshot("MKR56/26", "MKR48/26", "760010", "P", D("100"), D("40"), datetime(2026, 8, 2), "new")
        same = BackorderSnapshot("MKR56/26", "MKR56/26", "999999", "P", D("10"), D("1"), datetime(2026, 8, 2), "same")
        ledger = {}
        merge_latest_backorders(ledger, {old.key: old}, "MKR56/26")
        merge_latest_backorders(ledger, {new.key: new, same.key: same}, "MKR56/26")
        self.assertEqual(ledger[old.key].shipped_qty, D("40"))
        self.assertNotIn(same.key, ledger)

    def test_capacity_accepts_500_and_rejects_501(self):
        validate_capacity("MKR56/26", 500, 500, 500)
        with self.assertRaises(DataValidationError):
            validate_capacity("MKR56/26", 501, 0, 500)


if __name__ == "__main__":
    unittest.main()
