from __future__ import annotations

import queue
import os
import subprocess
import threading
import tkinter as tk
import sys
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from automation import run_automation
from config import (
    APP_CREDITS,
    BASE_DIR,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_LOG_DIR,
    DEFAULT_PARALLEL_SESSIONS,
    DEFAULT_STATE_PATH,
    MAX_PARALLEL_SESSIONS,
)
from template import create_input_template


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


COLORS = {
    "bg": "#f5f7fb",
    "surface": "#ffffff",
    "surface_alt": "#f8fafc",
    "border": "#e3eaf5",
    "text": "#172033",
    "muted": "#64748b",
    "primary": "#1554d1",
    "primary_hover": "#0f43aa",
    "primary_soft": "#e8f0ff",
    "success": "#0f8a52",
    "danger": "#dc2626",
    "danger_hover": "#b91c1c",
    "warning": "#b45309",
}


class KotraReportAppV2(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("KOTRA 보고서 자동 생성기")
        self.configure(fg_color=COLORS["bg"])

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = max(980, min(1180, screen_width - 120))
        height = max(680, min(840, screen_height - 120))
        self.geometry(f"{width}x{height}")
        self.minsize(940, 660)
        self.resizable(True, True)

        self.input_path = tk.StringVar(value=str(BASE_DIR / "input.xlsx"))
        self.download_dir = tk.StringVar(value=str(DEFAULT_DOWNLOAD_DIR))
        self.run_mode = tk.StringVar(value="전체 실행")
        self.background = tk.BooleanVar(value=False)
        self.use_session = tk.BooleanVar(value=False)
        self.parallel_sessions = tk.StringVar(value=str(DEFAULT_PARALLEL_SESSIONS))
        self.status = tk.StringVar(value="대기 중")
        self.progress = tk.StringVar(value="0 / 0")
        self.progress_percent = tk.DoubleVar(value=0)
        self.success = tk.StringVar(value="0건")
        self.failed = tk.StringVar(value="0건")
        self.current_row = tk.StringVar(value="0")
        self.total_rows = tk.StringVar(value="0")

        self.start_button: ctk.CTkButton | None = None
        self.stop_button: ctk.CTkButton | None = None
        self.force_stop_button: ctk.CTkButton | None = None
        self.status_badge: ctk.CTkLabel | None = None
        self.progress_bar: ctk.CTkProgressBar | None = None
        self.log_text: ctk.CTkTextbox | None = None
        self.parallel_options_frame: ctk.CTkFrame | None = None
        self.parallel_sessions_menu: ctk.CTkOptionMenu | None = None
        self.mode_buttons: dict[str, ctk.CTkButton] = {}

        self.stop_requested = False
        self.force_stop_requested = False
        self.parallel_options_enabled = False
        self.creator_click_count = 0
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_ui()
        self.after(200, self._poll_events)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        page = ctk.CTkScrollableFrame(self, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew", padx=24, pady=22)
        page.grid_columnconfigure(0, weight=1)

        self._build_header(page)
        self._build_file_panel(page)
        self._build_options_panel(page)
        self._build_action_panel(page)
        self._build_progress_panel(page)
        self._build_log_panel(page)

    def _build_header(self, parent: ctk.CTkFrame) -> None:
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text=APP_CREDITS["app_name"],
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=COLORS["text"],
        )
        title.grid(row=0, column=0, sticky="w")

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.grid(row=0, column=1, sticky="e", padx=(16, 0))

        self.status_badge = ctk.CTkLabel(
            right,
            textvariable=self.status,
            width=112,
            height=32,
            corner_radius=16,
            fg_color=COLORS["primary_soft"],
            text_color=COLORS["primary"],
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.status_badge.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            right,
            text="ⓘ",
            width=36,
            height=32,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=17, weight="bold"),
            command=self._show_about,
        ).pack(side="left")

        subtitle = ctk.CTkLabel(
            header,
            text="엑셀 데이터를 바탕으로 KOTRA 수출 시장 분석 보고서를 자동 생성하고 저장합니다.",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["muted"],
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(6, 0))

        version = ctk.CTkLabel(
            header,
            text=f"Version {APP_CREDITS['version']}",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["muted"],
        )
        version.grid(row=1, column=1, sticky="e", padx=(16, 0), pady=(6, 0))

    def _build_file_panel(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 1, "파일 설정")
        panel.grid_columnconfigure(1, weight=1)

        self._field_label(panel, "엑셀 파일").grid(row=0, column=0, sticky="w", padx=18, pady=(18, 8))
        ctk.CTkEntry(panel, textvariable=self.input_path, height=38, border_color=COLORS["border"]).grid(
            row=0, column=1, sticky="ew", pady=(18, 8)
        )
        ctk.CTkButton(
            panel,
            text="찾기",
            width=92,
            height=38,
            fg_color=COLORS["surface_alt"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._choose_input,
        ).grid(row=0, column=2, sticky="e", padx=18, pady=(18, 8))

        self._field_label(panel, "저장 폴더").grid(row=1, column=0, sticky="w", padx=18, pady=(8, 18))
        ctk.CTkEntry(panel, textvariable=self.download_dir, height=38, border_color=COLORS["border"]).grid(
            row=1, column=1, sticky="ew", pady=(8, 18)
        )
        ctk.CTkButton(
            panel,
            text="찾기",
            width=92,
            height=38,
            fg_color=COLORS["surface_alt"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._choose_download_dir,
        ).grid(row=1, column=2, sticky="e", padx=18, pady=(8, 18))

    def _build_options_panel(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 2, "실행 옵션")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_columnconfigure(1, weight=1)

        mode_box = ctk.CTkFrame(panel, fg_color="transparent")
        mode_box.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        mode_box.grid_columnconfigure(0, weight=1)
        self._field_label(mode_box, "실행 범위").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self._build_mode_selector(mode_box).grid(row=1, column=0, sticky="ew")

        switch_box = ctk.CTkFrame(
            panel,
            fg_color=COLORS["surface_alt"],
            border_width=1,
            border_color="#e2e8f0",
            corner_radius=8,
        )
        switch_box.grid(row=0, column=1, sticky="nsew", padx=18, pady=(18, 8))
        switch_box.grid_columnconfigure(0, weight=1)
        ctk.CTkSwitch(
            switch_box,
            text="백그라운드 실행",
            variable=self.background,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 10))
        ctk.CTkSwitch(
            switch_box,
            text="브라우저 세션 저장 사용",
            variable=self.use_session,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 14))

        self._build_parallel_options(panel)

        hint = ctk.CTkLabel(
            panel,
            text="실패 행 재시도는 logs/failed_rows.xlsx 기준입니다. 세션 저장은 로그인이나 쿠키 유지가 필요할 때만 사용하세요.",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["muted"],
            anchor="w",
        )
        hint.grid(row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(8, 18))

    def _build_parallel_options(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface_alt"],
            border_width=1,
            border_color="#dbeafe",
            corner_radius=8,
        )
        frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=18, pady=(8, 4))
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame,
            text="병렬 세션 수",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=16, pady=14)

        self.parallel_sessions_menu = ctk.CTkOptionMenu(
            frame,
            variable=self.parallel_sessions,
            values=[str(value) for value in range(1, MAX_PARALLEL_SESSIONS + 1)],
            width=96,
            height=34,
            fg_color=COLORS["surface"],
            button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
            text_color=COLORS["text"],
            dropdown_fg_color=COLORS["surface"],
        )
        self.parallel_sessions_menu.grid(row=0, column=1, sticky="e", padx=(12, 6), pady=14)

        ctk.CTkLabel(
            frame,
            text="개발자 옵션",
            text_color=COLORS["primary"],
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="e",
        ).grid(row=0, column=2, sticky="e", padx=(6, 16), pady=14)

        self.parallel_options_frame = frame
        frame.grid_remove()

    def _build_action_panel(self, parent: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(parent, fg_color="transparent")
        panel.grid(row=3, column=0, sticky="ew", pady=(0, 14))
        panel.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(panel, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w")

        self.start_button = ctk.CTkButton(
            left,
            text="▶ 시작",
            width=128,
            height=42,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=self._start,
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.start_button.pack(side="left")

        self.stop_button = ctk.CTkButton(
            left,
            text="■ 중지",
            width=112,
            height=42,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            text_color_disabled="#94a3b8",
            command=self._stop,
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=(8, 0))

        self.force_stop_button = ctk.CTkButton(
            left,
            text="× 강제종료",
            width=112,
            height=42,
            fg_color=COLORS["surface_alt"],
            hover_color=COLORS["surface_alt"],
            border_width=1,
            border_color=COLORS["border"],
            text_color="#94a3b8",
            text_color_disabled="#94a3b8",
            command=self._force_stop,
            state="disabled",
        )
        self.force_stop_button.pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            panel,
            text="+ 예시 템플릿 생성",
            width=148,
            height=42,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._create_template,
        ).grid(row=0, column=1, sticky="e")

        folders = ctk.CTkFrame(panel, fg_color="transparent")
        folders.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ctk.CTkButton(
            folders,
            text="↗ 다운로드 폴더 열기",
            width=148,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._open_download_dir,
        ).pack(side="left")

    def _build_progress_panel(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 4, "진행 현황")
        for column in range(4):
            panel.grid_columnconfigure(column, weight=1)

        self._metric_card(panel, 0, "현재 행", self.current_row, COLORS["primary"])
        self._metric_card(panel, 1, "전체 행", self.total_rows, COLORS["text"])
        self._metric_card(panel, 2, "성공", self.success, COLORS["success"])
        self._metric_card(panel, 3, "실패", self.failed, COLORS["danger"])

        progress_row = ctk.CTkFrame(panel, fg_color="transparent")
        progress_row.grid(row=1, column=0, columnspan=4, sticky="ew", padx=18, pady=(6, 14))
        progress_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(progress_row, textvariable=self.progress, text_color=COLORS["muted"], width=64).grid(
            row=0, column=0, sticky="w"
        )
        self.progress_bar = ctk.CTkProgressBar(
            progress_row,
            height=10,
            fg_color="#e2e8f0",
            progress_color=COLORS["primary"],
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=(12, 0))
        self.progress_bar.set(0)

        status_box = ctk.CTkFrame(panel, fg_color=COLORS["surface_alt"], corner_radius=8)
        status_box.grid(row=2, column=0, columnspan=4, sticky="ew", padx=18, pady=(0, 18))
        status_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            status_box,
            text="현재 상태",
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=12)
        ctk.CTkLabel(status_box, textvariable=self.status, text_color=COLORS["text"], anchor="w").grid(
            row=0, column=1, sticky="ew", padx=(0, 14), pady=12
        )

    def _build_log_panel(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 5, "실행 로그")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            panel,
            height=190,
            fg_color="#0f172a",
            text_color="#dbeafe",
            border_width=0,
            corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=13),
            wrap="word",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=18, pady=(18, 10))
        self.log_text.configure(state="disabled")
        self._configure_log_tags()

        log_actions = ctk.CTkFrame(panel, fg_color="transparent")
        log_actions.grid(row=1, column=0, sticky="e", padx=18, pady=(0, 18))
        ctk.CTkButton(
            log_actions,
            text="↗ 로그 폴더 열기",
            width=128,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._open_log_dir,
        ).pack(side="right")

    def _section(self, parent: ctk.CTkFrame, row: int, title: str) -> ctk.CTkFrame:
        wrapper = ctk.CTkFrame(parent, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=8)
        wrapper.grid(row=row, column=0, sticky="ew", pady=(0, 14))
        wrapper.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            wrapper,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 0))
        content = ctk.CTkFrame(wrapper, fg_color="transparent")
        content.grid(row=1, column=0, sticky="ew")
        return content

    def _field_label(self, parent: ctk.CTkFrame, text: str) -> ctk.CTkLabel:
        return ctk.CTkLabel(parent, text=text, text_color=COLORS["text"], font=ctk.CTkFont(size=14, weight="bold"))

    def _build_mode_selector(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        selector = ctk.CTkFrame(parent, fg_color="#eef3f9", corner_radius=10)
        for column in range(2):
            selector.grid_columnconfigure(column, weight=1)
        self.mode_buttons = {}
        modes = [("전체 실행", "전체 실행"), ("실패 행만 재시도", "실패 행만 재시도")]
        for column, (label, value) in enumerate(modes):
            button = ctk.CTkButton(
                selector,
                text=label,
                height=38,
                corner_radius=8,
                border_width=0,
                command=lambda selected=value: self._select_run_mode(selected),
                font=ctk.CTkFont(size=14, weight="bold"),
            )
            button.grid(row=0, column=column, sticky="ew", padx=(4 if column == 0 else 2, 4), pady=4)
            self.mode_buttons[value] = button
        self._refresh_run_mode_buttons()
        return selector

    def _select_run_mode(self, value: str) -> None:
        self.run_mode.set(value)
        self._refresh_run_mode_buttons()

    def _refresh_run_mode_buttons(self) -> None:
        for value, button in self.mode_buttons.items():
            selected = self.run_mode.get() == value
            button.configure(
                fg_color="#ffffff" if selected else "transparent",
                hover_color="#ffffff" if selected else "#e2eaf3",
                text_color=COLORS["primary"] if selected else COLORS["muted"],
            )

    def _metric_card(self, parent: ctk.CTkFrame, column: int, label: str, variable: tk.StringVar, color: str) -> None:
        card = ctk.CTkFrame(parent, fg_color=COLORS["surface_alt"], corner_radius=8)
        card.grid(row=0, column=column, sticky="ew", padx=(18 if column == 0 else 6, 18 if column == 3 else 6), pady=18)
        ctk.CTkLabel(card, text=label, text_color=COLORS["muted"], font=ctk.CTkFont(size=13)).pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(card, textvariable=variable, text_color=color, font=ctk.CTkFont(size=24, weight="bold")).pack(
            anchor="w", padx=14, pady=(0, 12)
        )

    def _configure_log_tags(self) -> None:
        if self.log_text is None:
            return
        self.log_text.tag_config("success", foreground="#86efac")
        self.log_text.tag_config("danger", foreground="#fca5a5")
        self.log_text.tag_config("warning", foreground="#fcd34d")
        self.log_text.tag_config("info", foreground="#bfdbfe")

    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.input_path.set(path)

    def _choose_download_dir(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.download_dir.set(path)

    def _create_template(self) -> None:
        path = create_input_template()
        messagebox.showinfo("완료", f"입력 템플릿을 생성했습니다.\n{path}")

    def _open_download_dir(self) -> None:
        self._open_folder(Path(self.download_dir.get()))

    def _open_log_dir(self) -> None:
        self._open_folder(DEFAULT_LOG_DIR)

    def _open_folder(self, path: Path) -> None:
        try:
            path.mkdir(parents=True, exist_ok=True)
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("폴더 열기 실패", f"폴더를 열 수 없습니다.\n{path}\n\n{exc}")

    def _show_about(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("프로그램 정보")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(fg_color=COLORS["bg"])

        body = ctk.CTkFrame(dialog, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=8)
        body.pack(fill="both", expand=True, padx=18, pady=18)
        body.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            body,
            text=APP_CREDITS["app_name"],
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 0))
        ctk.CTkLabel(
            body,
            text=f"Version {APP_CREDITS['version']}",
            font=ctk.CTkFont(size=13),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=22, pady=(4, 12))
        ctk.CTkLabel(
            body,
            text=APP_CREDITS["purpose"],
            font=ctk.CTkFont(size=14),
            text_color=COLORS["muted"],
            wraplength=460,
            justify="left",
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 18))

        info = ctk.CTkFrame(body, fg_color="transparent")
        info.grid(row=3, column=0, sticky="ew", padx=22, pady=(0, 18))
        info.grid_columnconfigure(1, weight=1)
        rows = [
            ("제작", APP_CREDITS["developed_by"]),
            ("소속/역할", APP_CREDITS["role"]),
            ("기술 스택", APP_CREDITS["tech_stack"]),
            ("개발 지원", APP_CREDITS["development_support"]),
            ("문의", APP_CREDITS["contact"]),
        ]
        for index, (label, value) in enumerate(rows):
            ctk.CTkLabel(
                info,
                text=label,
                width=92,
                font=ctk.CTkFont(size=13),
                text_color=COLORS["muted"],
                anchor="w",
            ).grid(row=index, column=0, sticky="w", pady=3)
            value_label = ctk.CTkLabel(
                info,
                text=value,
                font=ctk.CTkFont(size=13),
                text_color=COLORS["text"],
                wraplength=360,
                justify="left",
                anchor="w",
            )
            value_label.grid(row=index, column=1, sticky="ew", pady=3)
            if label == "제작":
                value_label.bind("<Button-1>", lambda _event: self._handle_parallel_easter_egg())

        divider = ctk.CTkFrame(body, height=1, fg_color=COLORS["border"])
        divider.grid(row=4, column=0, sticky="ew", padx=22, pady=(0, 14))

        ctk.CTkLabel(
            body,
            text=APP_CREDITS["disclaimer"],
            font=ctk.CTkFont(size=13),
            text_color=COLORS["muted"],
            wraplength=460,
            justify="left",
            anchor="w",
        ).grid(row=5, column=0, sticky="ew", padx=22, pady=(0, 8))
        ctk.CTkLabel(
            body,
            text=APP_CREDITS["copyright"],
            font=ctk.CTkFont(size=13),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=6, column=0, sticky="ew", padx=22, pady=(0, 18))

        ctk.CTkButton(
            body,
            text="확인",
            width=88,
            height=36,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_hover"],
            command=dialog.destroy,
        ).grid(row=7, column=0, sticky="e", padx=22, pady=(0, 22))

        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dialog.grab_set()
        dialog.lift()

    def _handle_parallel_easter_egg(self) -> None:
        if self.parallel_options_enabled:
            return

        self.creator_click_count += 1
        if self.creator_click_count < 5:
            return

        self.parallel_options_enabled = True
        self._show_parallel_options()
        self.status.set("병렬 처리 옵션 활성화됨")
        self._set_status_badge("warning")
        self._append_log("개발자 옵션: 병렬 처리 설정이 활성화되었습니다.", "warning")
        messagebox.showinfo("개발자 옵션", "병렬 처리 옵션이 활성화되었습니다.")

    def _show_parallel_options(self) -> None:
        if self.parallel_options_frame is not None:
            self.parallel_options_frame.grid()

    def _selected_parallel_sessions(self) -> int:
        if not self.parallel_options_enabled:
            return 1

        try:
            value = int(self.parallel_sessions.get())
        except ValueError:
            return DEFAULT_PARALLEL_SESSIONS

        return max(1, min(MAX_PARALLEL_SESSIONS, value))

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("실행 중", "이미 작업이 실행 중입니다.")
            return

        input_path = Path(self.input_path.get())
        if self._retry_failed_only() is False and not input_path.exists():
            messagebox.showerror("오류", f"엑셀 파일을 찾을 수 없습니다.\n{input_path}")
            return

        self.stop_requested = False
        self.force_stop_requested = False
        self.progress.set("0 / 0")
        self.progress_percent.set(0)
        self.current_row.set("0")
        self.total_rows.set("0")
        self.success.set("0건")
        self.failed.set("0건")
        self.status.set("시작 준비 중")
        self._set_status_badge("running")
        self._set_progress(0)
        self._clear_log()
        self._append_log("작업을 시작합니다.", "info")
        self._set_running_state(True)
        self.worker = threading.Thread(target=self._run_worker, daemon=True)
        self.worker.start()

    def _stop(self) -> None:
        self.stop_requested = True
        self.status.set("중지 요청됨: 현재 행 처리 후 멈춥니다.")
        self._set_status_badge("warning")
        self._append_log("중지 요청됨: 현재 행 처리 후 멈춥니다.", "warning")

    def _force_stop(self) -> None:
        self.force_stop_requested = True
        self.stop_requested = True
        self.status.set("강제종료 요청됨: 현재 작업을 즉시 중단합니다.")
        self._set_status_badge("danger")
        self._append_log("강제종료 요청됨: 현재 작업을 즉시 중단합니다.", "danger")

    def _retry_failed_only(self) -> bool:
        return self.run_mode.get() == "실패 행만 재시도"

    def _set_running_state(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state="disabled" if running else "normal")
        if self.stop_button is not None:
            self.stop_button.configure(state="normal" if running else "disabled")
        if self.force_stop_button is not None:
            if running:
                self.force_stop_button.configure(
                    state="normal",
                    fg_color=COLORS["danger"],
                    hover_color=COLORS["danger_hover"],
                    border_width=0,
                    text_color="#ffffff",
                )
            else:
                self.force_stop_button.configure(
                    state="disabled",
                    fg_color=COLORS["surface_alt"],
                    hover_color=COLORS["surface_alt"],
                    border_width=1,
                    border_color=COLORS["border"],
                    text_color="#94a3b8",
                )
        if self.parallel_sessions_menu is not None:
            self.parallel_sessions_menu.configure(state="disabled" if running else "normal")

    def _set_status_badge(self, state: str) -> None:
        if self.status_badge is None:
            return
        colors = {
            "idle": (COLORS["primary_soft"], COLORS["primary"]),
            "running": (COLORS["primary_soft"], COLORS["primary"]),
            "success": ("#dcfce7", COLORS["success"]),
            "warning": ("#fef3c7", COLORS["warning"]),
            "danger": ("#fee2e2", COLORS["danger"]),
        }
        fg, text = colors.get(state, colors["idle"])
        self.status_badge.configure(fg_color=fg, text_color=text)

    def _set_progress(self, percent: float) -> None:
        value = max(0.0, min(1.0, percent / 100))
        if self.progress_bar is not None:
            self.progress_bar.set(value)

    def _clear_log(self) -> None:
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _append_log(self, message: str, tag: str = "info") -> None:
        if self.log_text is None:
            return
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _run_worker(self) -> None:
        try:
            result = run_automation(
                input_excel_path=self.input_path.get(),
                download_dir=self.download_dir.get(),
                headless=self.background.get(),
                log_dir=DEFAULT_LOG_DIR,
                state_path=DEFAULT_STATE_PATH,
                use_storage_state=self.use_session.get(),
                save_storage_state=self.use_session.get(),
                retry_failed_only=self._retry_failed_only(),
                wait_for_manual_login=False,
                parallel_sessions=self._selected_parallel_sessions(),
                status_callback=lambda message: self.events.put(("status", message)),
                progress_callback=lambda data: self.events.put(("progress", data)),
                stop_requested=lambda: self.stop_requested,
                force_stop_requested=lambda: self.force_stop_requested,
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "status":
                    self.status.set(str(payload))
                    self._append_log(str(payload), self._log_tag(str(payload)))
                elif event == "progress":
                    self._handle_progress(payload)
                elif event == "done":
                    self._handle_done(payload)
                elif event == "error":
                    self.status.set("오류 발생")
                    self._set_status_badge("danger")
                    self._append_log(f"오류 발생: {payload}", "danger")
                    self._set_running_state(False)
                    messagebox.showerror("오류", str(payload))
        except queue.Empty:
            pass
        self.after(200, self._poll_events)

    def _handle_progress(self, payload: object) -> None:
        data = payload if isinstance(payload, dict) else {}
        current = int(data.get("current", 0) or 0)
        total = int(data.get("total", 0) or 0)
        success = int(data.get("success", 0) or 0)
        failed = int(data.get("failed", 0) or 0)
        percent = (current / total * 100) if total else 0

        self.current_row.set(str(current))
        self.total_rows.set(str(total))
        self.progress.set(f"{current} / {total}")
        self.progress_percent.set(percent)
        self._set_progress(percent)
        self.success.set(f"{success}건")
        self.failed.set(f"{failed}건")
        if data.get("status"):
            status = str(data["status"])
            self.status.set(status)
            self._append_log(status, self._log_tag(status))

    def _handle_done(self, payload: object) -> None:
        result = payload if isinstance(payload, dict) else {}
        failed = int(result.get("failed", 0) or 0)
        self.status.set("완료")
        self._set_status_badge("warning" if failed else "success")
        self._append_log("작업이 완료되었습니다.", "success" if failed == 0 else "warning")
        self._set_running_state(False)
        messagebox.showinfo(
            "완료",
            (
                f"전체 {result.get('total', 0)}건\n"
                f"성공 {result.get('success', 0)}건\n"
                f"실패 {result.get('failed', 0)}건"
            ),
        )

    def _log_tag(self, message: str) -> str:
        if "성공" in message or "완료" in message:
            return "success"
        if "실패" in message or "오류" in message or "강제종료" in message:
            return "danger"
        if "중지" in message or "재시도" in message:
            return "warning"
        return "info"


def run_gui() -> None:
    app = KotraReportAppV2()
    app.mainloop()
