from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .errors import ConfigurationError


REQUIRED_TARGETS = {
    "MKR48/26": "MKR48_26",
    "MKR48/26-1": "MKR48_26-1",
    "MKR49/26": "MKR49_26",
    "MKR49/26-1": "MKR49_26-1",
    "MKR56/26": "MKR56_26",
}


@dataclass(frozen=True)
class AppConfig:
    root: Path
    workbook: Path
    sender: str
    start_date: date
    max_messages: int
    data_start_row: int
    max_rows: int
    targets: dict[str, str]

    @property
    def data_last_row(self) -> int:
        return self.data_start_row + self.max_rows - 1

    @classmethod
    def load(cls, root: Path, filename: str = "config.json") -> "AppConfig":
        path = root / filename
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError as exc:
            raise ConfigurationError(f"설정 파일을 찾지 못했습니다: {path}") from exc
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigurationError(f"설정 파일을 읽을 수 없습니다: {exc}") from exc

        try:
            targets = {str(k): str(v) for k, v in raw["targets"].items()}
            config = cls(
                root=root,
                workbook=root / str(raw["workbook"]),
                sender=str(raw["sender"]).strip().lower(),
                start_date=date.fromisoformat(str(raw["start_date"])),
                max_messages=int(raw["max_messages"]),
                data_start_row=int(raw["data_start_row"]),
                max_rows=int(raw["max_rows"]),
                targets=targets,
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise ConfigurationError(f"설정 값이 올바르지 않습니다: {exc}") from exc

        if not config.sender or "@" not in config.sender:
            raise ConfigurationError("sender는 유효한 이메일 주소여야 합니다.")
        if not 1 <= config.max_messages <= 500:
            raise ConfigurationError("max_messages는 1~500이어야 합니다.")
        if config.data_start_row != 10 or config.max_rows != 500:
            raise ConfigurationError("V12.0 템플릿은 data_start_row=10, max_rows=500만 지원합니다.")
        if config.targets != REQUIRED_TARGETS:
            raise ConfigurationError("targets는 지정된 5개 MKR과 시트 매핑을 정확히 포함해야 합니다.")
        if not config.workbook.is_file():
            raise ConfigurationError(f"Excel 파일을 찾지 못했습니다: {config.workbook}")
        return config
