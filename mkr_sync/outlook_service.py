from __future__ import annotations

import base64
import json
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable

from .errors import DataValidationError, OfficeAutomationError
from .models import AttachmentEvent, CollectedMail, MailRecord, PreviewRequest
from .office import _load_com, _safe_filename, _sender_smtp, sha256_file
from .parsers import (
    attachment_kind,
    clean_text,
    current_message_body,
    find_all_mkrs,
    find_current_target_mkrs,
)


ProgressCallback = Callable[[int, str], None]


def _encode_folder(entry_id: str, store_id: str) -> str:
    payload = json.dumps({"entry": entry_id, "store": store_id}, separators=(",", ":"))
    return base64.urlsafe_b64encode(payload.encode("utf-8")).decode("ascii")


def _decode_folder(token: str) -> tuple[str, str]:
    try:
        raw = base64.urlsafe_b64decode(token.encode("ascii")).decode("utf-8")
        value = json.loads(raw)
        entry_id = str(value["entry"])
        store_id = str(value["store"])
    except Exception as exc:
        raise DataValidationError("Outlook 폴더 식별자가 올바르지 않습니다.") from exc
    if not entry_id or not store_id:
        raise DataValidationError("Outlook 폴더 식별자가 비어 있습니다.")
    return entry_id, store_id


def list_outlook_folders(max_depth: int = 3) -> list[dict[str, str]]:
    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception as exc:
            raise OfficeAutomationError(
                "실행 중인 Classic Outlook을 찾지 못했습니다. Classic Outlook을 먼저 여세요."
            ) from exc
        namespace = outlook.Session
        result: list[dict[str, str]] = []

        def visit(folder, label: str, depth: int) -> None:
            try:
                entry_id = clean_text(folder.EntryID)
                store_id = clean_text(folder.StoreID)
                if entry_id and store_id:
                    result.append(
                        {
                            "id": _encode_folder(entry_id, store_id),
                            "name": clean_text(folder.Name),
                            "path": label,
                        }
                    )
                if depth >= max_depth:
                    return
                folders = folder.Folders
                for index in range(1, int(folders.Count) + 1):
                    child = folders.Item(index)
                    visit(child, f"{label} / {clean_text(child.Name)}", depth + 1)
            except Exception:
                return

        roots = namespace.Folders
        for index in range(1, int(roots.Count) + 1):
            root = roots.Item(index)
            visit(root, clean_text(root.Name), 0)
        return result
    finally:
        pythoncom.CoUninitialize()


def collect_outlook_preview(
    request: PreviewRequest,
    run_attachment_dir: Path,
    logger: logging.Logger,
    cancel_event: threading.Event,
    progress: ProgressCallback | None = None,
) -> CollectedMail:
    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()
    result = CollectedMail()
    run_attachment_dir.mkdir(parents=True, exist_ok=True)
    targets = set(request.mkr_numbers)
    senders = {value.strip().lower() for value in request.senders}
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception as exc:
            raise OfficeAutomationError(
                "실행 중인 Classic Outlook을 찾지 못했습니다. Classic Outlook을 먼저 여세요."
            ) from exc
        namespace = outlook.Session
        entry_id, store_id = _decode_folder(request.outlook_folder_id)
        try:
            folder = namespace.GetFolderFromID(entry_id, store_id)
        except Exception as exc:
            raise DataValidationError("선택한 Outlook 폴더를 다시 찾지 못했습니다.") from exc

        folder_path = clean_text(getattr(folder, "FolderPath", "")) or clean_text(folder.Name)
        logger.info(
            "Outlook 검색 조건: 폴더=%s / 발신인=%s / 시작일=%s / 대상 MKR=%s / 최대 대상메일=%d",
            folder_path,
            ", ".join(sorted(senders)),
            request.start_date.isoformat(),
            ", ".join(request.mkr_numbers),
            request.max_messages,
        )

        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        count = int(items.Count)
        matched_messages = 0
        for index in range(1, count + 1):
            if cancel_event.is_set():
                break
            mail = items.Item(index)
            result.scanned += 1
            try:
                if int(mail.Class) != 43:
                    continue
                received = mail.ReceivedTime
                if not isinstance(received, datetime):
                    received = datetime.fromisoformat(str(received))
                if received.date() < request.start_date:
                    break
                if _sender_smtp(mail) not in senders:
                    continue
                subject = clean_text(mail.Subject)
                body = current_message_body(mail.Body)
                target_keys = find_current_target_mkrs(subject, body, targets)
                if not target_keys:
                    continue
                if matched_messages >= request.max_messages:
                    logger.info("Outlook 대상 메일 상한 %d건에 도달했습니다.", request.max_messages)
                    break
                matched_messages += 1
                mail_entry_id = clean_text(mail.EntryID) or f"mail-{index}-{received:%Y%m%d%H%M%S}"

                attachment_descriptors: list[tuple[object, str, str, bool, list[str]]] = []
                for attachment_index in range(1, int(mail.Attachments.Count) + 1):
                    attachment = mail.Attachments.Item(attachment_index)
                    filename = clean_text(attachment.FileName)
                    kind, revision = attachment_kind(filename, subject)
                    if kind is not None:
                        attachment_descriptors.append(
                            (attachment, filename, kind, revision, find_all_mkrs(filename))
                        )

                for target in target_keys:
                    result.records.append(MailRecord(mail_entry_id, target, received, subject, body))
                    for attachment_index, descriptor in enumerate(attachment_descriptors, start=1):
                        attachment, filename, kind, revision, file_mkrs = descriptor
                        if file_mkrs and target not in file_mkrs:
                            continue
                        destination = run_attachment_dir / (
                            f"{received:%Y%m%d_%H%M%S}_{index:03d}_{attachment_index:02d}_"
                            f"{_safe_filename(target)}_{_safe_filename(filename)}"
                        )
                        attachment.SaveAsFile(str(destination))
                        result.attachments.append(
                            AttachmentEvent(
                                entry_id=mail_entry_id,
                                target_mkr=target,
                                received_at=received,
                                subject=subject,
                                body=body,
                                kind=kind,
                                is_revision=revision,
                                original_name=filename,
                                path=destination,
                                sha256=sha256_file(destination),
                            )
                        )
                        logger.info("첨부 저장: %s / %s / %s", target, kind, filename)
            finally:
                del mail
            if progress and (index == 1 or index % 10 == 0 or index == count):
                progress(
                    min(99, int(index / max(count, 1) * 100)),
                    f"Outlook 메일 {index}/{count}건 확인 · 대상 {matched_messages}건",
                )

        if cancel_event.is_set():
            return result
        if not result.records:
            raise DataValidationError("조건에 맞는 대상 MKR 메일을 찾지 못했습니다.")
        logger.info(
            "Outlook 검색 완료: 확인 %d건 / 대상 메일 %d건 / 첨부 %d개",
            result.scanned,
            matched_messages,
            len(result.attachments),
        )
        return result
    finally:
        pythoncom.CoUninitialize()
