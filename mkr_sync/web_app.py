from __future__ import annotations

import ctypes
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from .web_api import WebApi


def application_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative


def configure_logging(root: Path) -> tuple[logging.Logger, Path]:
    log_dir = root / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / f"MKR_SYNC_{datetime.now():%Y%m%d_%H%M%S}.log"
    logger = logging.getLogger("mkr_sync.web")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.FileHandler(path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(message)s"))
    logger.addHandler(handler)
    return logger, path


def _message_box(message: str, title: str = "MKR SYNC") -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except Exception:
        pass


def run_web_app() -> int:
    root = application_root()
    logger, log_path = configure_logging(root)
    try:
        import webview  # type: ignore

        html_path = resource_path("mkr_sync/web/index.html")
        html = html_path.read_text(encoding="utf-8")
        api = WebApi(root, logger)
        window = webview.create_window(
            "MKR SYNC",
            html=html,
            js_api=api,
            width=1180,
            height=820,
            min_size=(900, 650),
            resizable=True,
            text_select=True,
        )
        if window is None:
            raise RuntimeError("pywebview 창을 만들지 못했습니다.")
        window.events.loaded += lambda: logger.info("HTML 화면 로드 완료")
        window.events.closed += lambda: logger.info("MKR SYNC 창 종료")
        os.environ["PYWEBVIEW_GUI"] = "edgechromium"
        logger.info("Edge WebView2 렌더러로 HTML 내장 창을 시작합니다.")
        webview.start(gui="edgechromium", debug=False, private_mode=True)
        return 0
    except Exception as exc:
        logger.exception("HTML 내장 앱을 시작하지 못했습니다.")
        _message_box(
            "MKR SYNC를 시작하지 못했습니다.\n\n"
            "Microsoft Edge WebView2 Runtime, Python 패키지 또는 실행 파일 구성을 확인해 주세요.\n\n"
            f"오류: {exc}\n로그: {log_path}"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(run_web_app())
