from __future__ import annotations

import argparse
import logging
import sys
import unittest
from datetime import datetime
from pathlib import Path

from .config import AppConfig
from .errors import (
    ConfigurationError,
    DataValidationError,
    MkrSyncError,
    OfficeAutomationError,
    QuantityRuleError,
    WorkbookCommitError,
)
from .office import WorkbookTransaction, assert_workbook_available, collect_outlook
from .reporting import write_audit_report
from .workflow import build_results


ROOT = Path(__file__).resolve().parent.parent


def _configure_logging() -> tuple[logging.Logger, Path]:
    log_dir = ROOT / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / f"MKR_SYNC_{datetime.now():%Y%m%d_%H%M%S}.log"
    logger = logging.getLogger("mkr_sync")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler = logging.FileHandler(path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger, path


def run_self_tests() -> int:
    suite = unittest.defaultTestLoader.discover(str(ROOT / "tests"), pattern="test_*.py")
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() else 1


def run_sync(dry_run: bool) -> int:
    logger, log_path = _configure_logging()
    transaction = None
    try:
        config = AppConfig.load(ROOT)
        if dry_run:
            if not config.workbook.is_file():
                raise ConfigurationError(f"Excel 파일을 찾지 못했습니다: {config.workbook}")
        else:
            transaction = WorkbookTransaction(config, logger)
            transaction.prepare()

        attachment_dir = ROOT / "MKR_Attachments" / f"Run_{datetime.now():%Y%m%d_%H%M%S}"
        collected = collect_outlook(config, attachment_dir, logger)
        results = build_results(config, collected, logger)
        report = write_audit_report(ROOT, results, dry_run=dry_run)
        logger.info("감사 보고서 생성: %s", report)

        if dry_run:
            logger.info("DRY RUN 완료: Excel 원본은 변경하지 않았습니다.")
            print(f"[OK] 미리보기 완료: {report}")
            return 0

        assert transaction is not None
        transaction.write_pending(results)
        transaction.validate_pending(results)
        transaction.commit()
        logger.info(
            "완료: 검색 %d건 / 반영 시트 %d개 / 보고서 %s",
            collected.scanned,
            len(results),
            report,
        )
        print(f"[OK] 동기화 완료: {config.workbook}")
        return 0
    except QuantityRuleError as exc:
        logger.error("수량 판정 중단: %s", exc)
        return 4
    except ConfigurationError as exc:
        logger.error("설정 오류: %s", exc)
        return 2
    except OfficeAutomationError as exc:
        logger.error("Office 연동 오류: %s", exc)
        return 3
    except DataValidationError as exc:
        logger.error("자료 검증 오류: %s", exc)
        return 4
    except WorkbookCommitError as exc:
        logger.error("Excel 저장 오류: %s", exc)
        return 5
    except MkrSyncError as exc:
        logger.error("실행 중단: %s", exc)
        return 6
    except Exception:
        logger.exception("예상하지 못한 오류로 중단했습니다.")
        return 9
    finally:
        if transaction is not None:
            transaction.cleanup_pending()
        print(f"로그: {log_path}")


def run_gui() -> int:
    from .web_app import run_web_app

    return run_web_app()


def run_tk_gui() -> int:
    from .gui import run_gui as start_gui

    return start_gui()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="MKR FlowSync V12.0 Python Edition")
    parser.add_argument("command", choices=("self-test", "dry-run", "sync", "gui", "tk-gui"))
    args = parser.parse_args(argv)
    if args.command == "self-test":
        return run_self_tests()
    if args.command == "gui":
        return run_gui()
    if args.command == "tk-gui":
        return run_tk_gui()
    return run_sync(dry_run=args.command == "dry-run")


if __name__ == "__main__":
    raise SystemExit(main())
