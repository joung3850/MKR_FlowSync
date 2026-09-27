from __future__ import annotations

import logging
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .errors import MkrSyncError
from .mkr_parser import extract_mkr_numbers, to_excel_sheet_name
from .sheet_creator import create_mkr_sheets


ROOT = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent.parent
)


def _configure_gui_logging() -> logging.Logger:
    log_dir = ROOT / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mkr_sync.gui")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.FileHandler(log_dir / "MKR_GUI.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(message)s"))
    logger.addHandler(handler)
    return logger


class MkrSyncApp:
    def __init__(self, root: tk.Tk, logger: logging.Logger):
        self.root = root
        self.logger = logger
        self.workbook_path: Path | None = None
        self.file_var = tk.StringVar(value="선택된 파일 없음")
        self.status_var = tk.StringVar()
        self._build()

    def _build(self) -> None:
        self.root.title("MKR SYNC")
        self.root.geometry("760x650")
        self.root.minsize(620, 520)
        self.root.configure(bg="#f4f7fb")

        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Malgun Gothic", 22, "bold"), foreground="#10233f")
        style.configure("Section.TLabel", font=("Malgun Gothic", 12, "bold"), foreground="#10233f")
        style.configure("Primary.TButton", font=("Malgun Gothic", 12, "bold"), padding=9)

        outer = ttk.Frame(self.root, padding=28)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="▣  MKR SYNC", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="설명문에서 MKR 번호를 찾아 Excel 시트를 생성합니다.").pack(anchor="w", pady=(6, 24))

        file_frame = ttk.LabelFrame(outer, text="엑셀 파일", padding=14)
        file_frame.pack(fill="x", pady=(0, 16))
        ttk.Label(file_frame, textvariable=self.file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="파일 선택", command=self.choose_file).pack(side="right", padx=(12, 0))

        input_frame = ttk.LabelFrame(outer, text="MKR 번호 또는 설명문 입력", padding=14)
        input_frame.pack(fill="both", expand=True, pady=(0, 16))
        self.input_text = tk.Text(input_frame, height=10, wrap="word", font=("Malgun Gothic", 11), relief="solid", borderwidth=1)
        self.input_text.pack(fill="both", expand=True)
        self.input_text.bind("<KeyRelease>", lambda _event: self.preview_numbers())

        preview_frame = ttk.LabelFrame(outer, text="추출된 MKR 번호", padding=10)
        preview_frame.pack(fill="x", pady=(0, 16))
        self.preview = tk.Listbox(preview_frame, height=4, font=("Consolas", 10), activestyle="none")
        self.preview.pack(fill="x")

        ttk.Button(outer, text="시트 생성", style="Primary.TButton", command=self.create_sheets).pack(fill="x")
        ttk.Label(outer, textvariable=self.status_var, foreground="#08794c", wraplength=680).pack(pady=(14, 0))

    def choose_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            self.workbook_path = Path(selected)
            self.file_var.set(str(self.workbook_path))
            self.status_var.set("")

    def preview_numbers(self) -> list[str]:
        numbers = extract_mkr_numbers(self.input_text.get("1.0", "end-1c"))
        self.preview.delete(0, tk.END)
        for number in numbers:
            self.preview.insert(tk.END, f"{number}  →  {to_excel_sheet_name(number)}")
        return numbers

    def create_sheets(self) -> None:
        numbers = self.preview_numbers()
        if self.workbook_path is None:
            self.show_error("엑셀 파일을 먼저 선택해 주세요.")
            return
        if not numbers:
            self.show_error("입력문에서 MKR 번호를 찾지 못했습니다.")
            return
        try:
            self.logger.info("시트 생성 대상 Excel: %s", self.workbook_path)
            result = create_mkr_sheets(
                self.workbook_path,
                numbers,
                template_workbook=ROOT / "MKR_TEMPLATE.xlsx",
                logger=self.logger,
                backup_dir=ROOT / "Backup",
            )
        except (MkrSyncError, ValueError) as exc:
            self.logger.exception("GUI 시트 생성 실패")
            self.show_error(str(exc))
            return
        created = result["created"]
        existing = result["existing"]
        summary = [*(f"생성: {item}" for item in created), *(f"기존: {item}" for item in existing)]
        self.status_var.set(f"생성 {len(created)}개 / 기존 정상 {len(existing)}개")
        messagebox.showinfo("MKR SYNC", "시트 준비가 완료되었습니다.\n\n" + "\n".join(summary))

    def show_error(self, message: str) -> None:
        self.status_var.set(message)
        messagebox.showerror("MKR SYNC 오류", message)


def run_gui() -> int:
    logger = _configure_gui_logging()
    root = tk.Tk()
    MkrSyncApp(root, logger)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(run_gui())
