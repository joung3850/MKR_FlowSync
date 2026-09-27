import logging
import time
import unittest

from mkr_sync.jobs import JobManager


class JobManagerTests(unittest.TestCase):
    def wait(self, manager, job_id):
        for _ in range(100):
            status = manager.status(job_id)
            if status["state"] not in {"queued", "running"}:
                return status
            time.sleep(0.01)
        self.fail("job did not complete")

    def test_completed_job_is_json_safe(self):
        manager = JobManager(logging.getLogger("job-test"))
        job_id = manager.start(
            "test",
            lambda cancel, progress: {
                "value": 7,
                "warnings": [
                    {"level": "warning", "code": "W", "message": "warn", "mkr": "MKR1/26"}
                ],
                "errors": [],
            },
        )
        status = self.wait(manager, job_id)
        self.assertEqual(status["state"], "completed")
        self.assertEqual(status["result"]["value"], 7)
        self.assertEqual(status["warnings"][0]["code"], "W")

    def test_rejects_overlapping_job(self):
        manager = JobManager(logging.getLogger("job-test"))

        def slow(cancel, progress):
            time.sleep(0.08)
            return {}

        manager.start("test", slow)
        with self.assertRaises(RuntimeError):
            manager.start("test", slow)


if __name__ == "__main__":
    unittest.main()
