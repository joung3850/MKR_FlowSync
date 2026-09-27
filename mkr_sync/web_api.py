from __future__ import annotations

import json
import logging
import os
import re
from datetime import date
from pathlib import Path

from .errors import DataValidationError
from .apply_service import apply_preview_results
from .jobs import JobManager
from .models import PreviewRequest
from .mkr_parser import extract_mkr_numbers, to_excel_sheet_name
from .outlook_service import list_outlook_folders
from .preview_service import run_preview
from .sheet_creator import _disk_sheet_names, create_mkr_sheets


EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
EXCEL_TYPES = ("Excel files (*.xlsx;*.xlsm)", "All files (*.*)")


class WebApi:
    def __init__(self, root: Path, logger: logging.Logger):
        self.root = Path(root).resolve()
        self.logger = logger
        self.settings_path = self.root / "settings.json"
        self.jobs = JobManager(logger)
        self.target_workbook: Path | None = None
        self.template_workbook: Path | None = self._resolve_template(False)
        self._applied_preview_jobs: set[str] = set()

    def _settings(self) -> dict:
        try:
            value = json.loads(self.settings_path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else {}
        except (FileNotFoundError, OSError, json.JSONDecodeError):
            return {}

    def _save_settings(self, updates: dict) -> None:
        values = self._settings()
        values.update(updates)
        pending = self.settings_path.with_suffix(".json.pending")
        pending.write_text(
            json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        os.replace(pending, self.settings_path)

    def _resolve_template(self, prompt: bool) -> Path | None:
        default = self.root / "MKR_TEMPLATE.xlsx"
        if default.is_file():
            return default.resolve()
        saved = str(self._settings().get("template_workbook", "")).strip()
        if saved and Path(saved).is_file():
            return Path(saved).resolve()
        if prompt:
            selected = self._choose_file("MKR_TEMPLATE.xlsx 선택")
            if selected:
                self.template_workbook = selected
                self._save_settings({"template_workbook": str(selected)})
                return selected
        return None

    def _choose_file(self, title: str) -> Path | None:
        try:
            import webview  # type: ignore

            window = webview.active_window()
            result = window.create_file_dialog(
                webview.FileDialog.OPEN,
                allow_multiple=False,
                file_types=EXCEL_TYPES,
            )
        except Exception as exc:
            self.logger.exception("파일 선택창 오류")
            raise RuntimeError(f"파일 선택창을 열지 못했습니다: {exc}") from exc
        if not result:
            return None
        value = result[0] if isinstance(result, (list, tuple)) else result
        path = Path(str(value)).resolve()
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise DataValidationError(f"올바른 Excel 파일이 아닙니다: {path}")
        return path

    @staticmethod
    def _ok(**values) -> dict:
        return {"ok": True, **values}

    @staticmethod
    def _error(exc: Exception) -> dict:
        return {"ok": False, "error": str(exc), "error_type": type(exc).__name__}

    def get_startup_state(self) -> dict:
        return self._ok(
            target_workbook=str(self.target_workbook) if self.target_workbook else "",
            template_workbook=str(self.template_workbook) if self.template_workbook else "",
            template_available=bool(self.template_workbook and self.template_workbook.is_file()),
            root=str(self.root),
            preview_only=False,
        )

    def choose_target_workbook(self) -> dict:
        try:
            selected = self._choose_file("대상 Excel 선택")
            if selected is None:
                return self._ok(cancelled=True, path="")
            self.target_workbook = selected
            return self._ok(cancelled=False, path=str(selected), name=selected.name)
        except Exception as exc:
            return self._error(exc)

    def choose_template_workbook(self) -> dict:
        try:
            selected = self._choose_file("MKR 템플릿 Excel 선택")
            if selected is None:
                return self._ok(cancelled=True, path="")
            if self.target_workbook and os.path.normcase(str(selected)) == os.path.normcase(str(self.target_workbook)):
                raise DataValidationError("대상 Excel과 템플릿 Excel은 서로 달라야 합니다.")
            self.template_workbook = selected
            self._save_settings({"template_workbook": str(selected)})
            return self._ok(cancelled=False, path=str(selected), name=selected.name)
        except Exception as exc:
            return self._error(exc)

    def extract_mkr_numbers(self, text: str) -> dict:
        try:
            numbers = extract_mkr_numbers(str(text or ""))
            return self._ok(
                numbers=numbers,
                sheets=[to_excel_sheet_name(item) for item in numbers],
                count=len(numbers),
            )
        except Exception as exc:
            return self._error(exc)

    def create_mkr_sheets(self, request: dict) -> dict:
        try:
            if self.target_workbook is None:
                raise DataValidationError("대상 Excel을 먼저 선택해 주세요.")
            numbers = extract_mkr_numbers(str(request.get("text", "")))
            if not numbers:
                raise DataValidationError("입력문에서 MKR 번호를 찾지 못했습니다.")
            disk_names = _disk_sheet_names(self.target_workbook) or set()
            requested_sheets = {to_excel_sheet_name(item) for item in numbers}
            template = self._resolve_template(True) if requested_sheets.difference(disk_names) else None

            def worker(cancel_event, progress):
                # Sheet commit is deliberately not interrupted after it begins.
                result = create_mkr_sheets(
                    self.target_workbook,
                    numbers,
                    template_workbook=template,
                    logger=self.logger,
                    backup_dir=self.root / "Backup",
                    progress=lambda value, message: progress(value, "sheet", message),
                )
                return {**result, "mkr_numbers": numbers}

            job_id = self.jobs.start("sheet", worker)
            return self._ok(job_id=job_id, numbers=numbers)
        except Exception as exc:
            return self._error(exc)

    def list_outlook_folders(self) -> list[dict]:
        try:
            return [{"ok": True, **item} for item in list_outlook_folders()]
        except Exception as exc:
            return [{"ok": False, "error": str(exc), "error_type": type(exc).__name__}]

    def start_preview(self, request: dict) -> dict:
        try:
            if self.target_workbook is None:
                raise DataValidationError("대상 Excel을 먼저 선택해 주세요.")
            numbers = extract_mkr_numbers(str(request.get("text", "")))
            if not numbers:
                raise DataValidationError("분석할 MKR 번호를 찾지 못했습니다.")
            raw_senders = request.get("senders", [])
            if isinstance(raw_senders, str):
                raw_senders = re.split(r"[,;\n]+", raw_senders)
            senders = []
            for value in raw_senders:
                sender = str(value).strip().lower()
                if sender and sender not in senders:
                    senders.append(sender)
            if not senders or any(not EMAIL_PATTERN.fullmatch(value) for value in senders):
                raise DataValidationError("발신인 이메일 주소를 올바르게 입력해 주세요.")
            if len(senders) > 10:
                raise DataValidationError("발신인은 최대 10명까지 입력할 수 있습니다.")
            start_date = date.fromisoformat(str(request.get("start_date", "")))
            max_messages = int(request.get("max_messages", 500))
            if not 1 <= max_messages <= 500:
                raise DataValidationError("최대 메일 수는 1~500이어야 합니다.")
            folder_id = str(request.get("outlook_folder_id", "")).strip()
            if not folder_id:
                raise DataValidationError("Outlook 폴더를 선택해 주세요.")
            preview_request = PreviewRequest(
                target_workbook=self.target_workbook,
                mkr_numbers=numbers,
                senders=senders,
                outlook_folder_id=folder_id,
                start_date=start_date,
                max_messages=max_messages,
            )

            def worker(cancel_event, progress):
                return run_preview(self.root, preview_request, self.logger, cancel_event, progress)

            job_id = self.jobs.start("outlook", worker)
            return self._ok(job_id=job_id, numbers=numbers)
        except Exception as exc:
            return self._error(exc)

    def get_job_status(self, job_id: str) -> dict:
        try:
            return self._ok(status=self.jobs.status(str(job_id)))
        except Exception as exc:
            return self._error(exc)

    def apply_preview(self, request: dict) -> dict:
        try:
            if self.target_workbook is None:
                raise DataValidationError("대상 Excel을 먼저 선택해 주세요.")
            preview_job_id = str(request.get("preview_job_id", "")).strip()
            if not preview_job_id:
                raise DataValidationError("적용할 미리보기 작업을 찾지 못했습니다.")
            if preview_job_id in self._applied_preview_jobs:
                raise DataValidationError("이 미리보기 결과는 이미 Excel에 적용되었습니다.")
            preview_status = self.jobs.status(preview_job_id)
            if preview_status.get("state") != "completed":
                raise DataValidationError("완료된 미리보기 결과만 Excel에 적용할 수 있습니다.")
            preview_result = preview_status.get("result")
            if not isinstance(preview_result, dict):
                raise DataValidationError("미리보기 결과를 불러오지 못했습니다.")
            if preview_result.get("errors"):
                raise DataValidationError("미리보기 오류를 해결한 뒤 다시 분석해 주세요.")
            target_workbook = self.target_workbook

            def worker(cancel_event, progress):
                result = apply_preview_results(
                    self.root,
                    target_workbook,
                    preview_result,
                    self.logger,
                    cancel_event,
                    progress,
                )
                self._applied_preview_jobs.add(preview_job_id)
                return result

            job_id = self.jobs.start("apply", worker)
            return self._ok(job_id=job_id, preview_job_id=preview_job_id)
        except Exception as exc:
            return self._error(exc)

    def cancel_job(self, job_id: str) -> dict:
        try:
            return self._ok(status=self.jobs.cancel(str(job_id)))
        except Exception as exc:
            return self._error(exc)

    def open_report_folder(self) -> dict:
        try:
            path = self.root / "Reports"
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
            return self._ok(path=str(path))
        except Exception as exc:
            return self._error(exc)
