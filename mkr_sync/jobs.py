from __future__ import annotations

import logging
import threading
import uuid
from collections.abc import Callable

from .models import AnalysisIssue, JobStatus


class JobCancelled(Exception):
    pass


ProgressCallback = Callable[[int, str, str], None]
Worker = Callable[[threading.Event, ProgressCallback], dict]


class JobManager:
    """Run at most one Office job and expose JSON-safe status snapshots."""

    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self._lock = threading.RLock()
        self._jobs: dict[str, JobStatus] = {}
        self._cancel_events: dict[str, threading.Event] = {}
        self._active_job_id: str | None = None

    def start(self, phase: str, worker: Worker) -> str:
        with self._lock:
            if self._active_job_id:
                active = self._jobs[self._active_job_id]
                if active.state in {"queued", "running"}:
                    raise RuntimeError("다른 작업이 실행 중입니다. 완료하거나 취소한 뒤 다시 실행하세요.")
            job_id = uuid.uuid4().hex
            self._jobs[job_id] = JobStatus(
                job_id=job_id,
                state="queued",
                phase=phase,
                progress=0,
                message="작업을 준비하고 있습니다.",
            )
            cancel_event = threading.Event()
            self._cancel_events[job_id] = cancel_event
            self._active_job_id = job_id
            thread = threading.Thread(
                target=self._run,
                args=(job_id, worker, cancel_event),
                daemon=True,
                name=f"mkr-job-{job_id[:8]}",
            )
            thread.start()
            return job_id

    def _run(self, job_id: str, worker: Worker, cancel_event: threading.Event) -> None:
        self._update(job_id, state="running", progress=1, message="작업을 시작했습니다.")

        def progress(value: int, phase: str, message: str) -> None:
            self._update(
                job_id,
                progress=max(0, min(99, int(value))),
                phase=phase,
                message=message,
            )

        try:
            result = worker(cancel_event, progress)
            if cancel_event.is_set():
                raise JobCancelled("사용자 요청으로 작업을 취소했습니다.")
            warnings = [AnalysisIssue(**item) for item in result.get("warnings", [])]
            errors = [AnalysisIssue(**item) for item in result.get("errors", [])]
            self._update(
                job_id,
                state="completed",
                progress=100,
                message="작업이 완료되었습니다.",
                result=result,
                warnings=warnings,
                errors=errors,
            )
        except JobCancelled as exc:
            self._update(job_id, state="cancelled", message=str(exc))
        except Exception as exc:
            self.logger.exception("백그라운드 작업 실패: %s", job_id)
            self._update(
                job_id,
                state="failed",
                message=str(exc),
                errors=[AnalysisIssue("error", type(exc).__name__, str(exc))],
            )
        finally:
            with self._lock:
                if self._active_job_id == job_id:
                    self._active_job_id = None

    def _update(self, job_id: str, **changes) -> None:
        with self._lock:
            status = self._jobs[job_id]
            for key, value in changes.items():
                setattr(status, key, value)

    def status(self, job_id: str) -> dict:
        with self._lock:
            if job_id not in self._jobs:
                raise KeyError("작업 ID를 찾지 못했습니다.")
            return self._jobs[job_id].as_dict()

    def cancel(self, job_id: str) -> dict:
        with self._lock:
            if job_id not in self._jobs:
                raise KeyError("작업 ID를 찾지 못했습니다.")
            status = self._jobs[job_id]
            if status.state not in {"queued", "running"}:
                return status.as_dict()
            self._cancel_events[job_id].set()
            status.message = "현재 처리 단위가 끝나면 취소합니다."
            return status.as_dict()
