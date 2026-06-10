from __future__ import annotations

import queue
import os
import subprocess
import threading
import tkinter as tk
import tkinter.font as tkfont
import sys
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from automation import FILENAME_PATTERN_TOKEN_LABELS, render_filename_pattern, run_automation
from config import (
    APP_CREDITS,
    BASE_DIR,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_LOG_DIR,
    DEFAULT_PARALLEL_SESSIONS,
    DEFAULT_ROW_RETRY_COUNT,
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

DEFAULT_FILENAME_PATTERN = ""
DEFAULT_FILENAME_PARTS: list[dict[str, str]] = []
FILENAME_TOKEN_LABEL_BY_TOKEN = {token: label for label, token in FILENAME_PATTERN_TOKEN_LABELS}
FILENAME_TEXT_SHORTCUTS = [("_", "_"), ("-", "-"), (".", "."), ("공백", " "), ("(", "("), (")", ")"), ("[", "["), ("]", "]")]
FILENAME_PRESETS = [
    (
        "연번_HS_품명",
        [
            {"type": "token", "value": "row_index"},
            {"type": "text", "value": "_"},
            {"type": "token", "value": "hs_code"},
            {"type": "text", "value": "_"},
            {"type": "token", "value": "product_name"},
        ],
    ),
    (
        "국가_품명[HS]",
        [
            {"type": "token", "value": "target_country"},
            {"type": "text", "value": "_"},
            {"type": "token", "value": "product_name"},
            {"type": "text", "value": "["},
            {"type": "token", "value": "hs_code"},
            {"type": "text", "value": "]"},
        ],
    ),
    (
        "사이트기본_생성일시",
        [
            {"type": "token", "value": "site_filename"},
            {"type": "text", "value": "_"},
            {"type": "token", "value": "datetime"},
        ],
    ),
]
FILENAME_PREVIEW_ROW = {
    "row_index": 1,
    "hs_code": "330499",
    "product_name": "스킨케어",
    "target_country": "베트남",
    "excluded_countries": "중국",
    "report_mode": "direct",
}


class HoverTooltip:
    active: HoverTooltip | None = None
    scheduled: HoverTooltip | None = None

    def __init__(self, widget: tk.Widget, text: str, delay_ms: int = 700, max_width: int = 720) -> None:
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self.max_width = max_width
        self.after_id: str | None = None
        self.tooltip: tk.Toplevel | None = None

        self._bind_widget(widget)
        for attr in ("_canvas", "_text_label", "_image_label", "_switch", "_button", "_label"):
            child = getattr(widget, attr, None)
            if child is not None:
                self._bind_widget(child)

    def _bind_widget(self, widget: tk.Widget) -> None:
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<FocusIn>", self._schedule, add="+")
        widget.bind("<FocusOut>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event: tk.Event | None = None) -> None:
        if HoverTooltip.scheduled is not None and HoverTooltip.scheduled is not self:
            HoverTooltip.scheduled._cancel()
        if HoverTooltip.active is not None and HoverTooltip.active is not self:
            HoverTooltip.active._hide_now()
        self._cancel()
        HoverTooltip.scheduled = self
        self.after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self) -> None:
        if self.after_id is not None:
            self.widget.after_cancel(self.after_id)
            self.after_id = None
        if HoverTooltip.scheduled is self:
            HoverTooltip.scheduled = None

    def _show(self) -> None:
        self.after_id = None
        if HoverTooltip.scheduled is self:
            HoverTooltip.scheduled = None
        if self.tooltip is not None or not self.text:
            return

        if HoverTooltip.active is not None and HoverTooltip.active is not self:
            HoverTooltip.active._hide_now()

        root = self.widget.winfo_toplevel()
        if not root.winfo_viewable():
            return

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        if sys.platform == "darwin":
            self.tooltip.configure(bg="systemTransparent")
            self.tooltip.wm_attributes("-transparent", True)
        else:
            transparent_color = "#010203"
            self.tooltip.configure(bg=transparent_color)
            try:
                self.tooltip.wm_attributes("-transparentcolor", transparent_color)
            except tk.TclError:
                self.tooltip.configure(bg=COLORS["text"])
        font_family = self._font_family()
        font_size = 12 if sys.platform.startswith("win") else 11
        measure_font = tkfont.nametofont("TkDefaultFont").copy()
        measure_font.configure(family=font_family, size=font_size)
        label_font = ctk.CTkFont(family=font_family, size=font_size)
        padx = 12
        screen_width = self.widget.winfo_screenwidth()
        screen_height = self.widget.winfo_screenheight()
        max_label_width = max(180, min(self.max_width, screen_width - 40 - (padx * 2)))
        text_width = measure_font.measure(self.text)
        wraplength = min(max_label_width, max(1, text_width))

        body = ctk.CTkFrame(
            self.tooltip,
            fg_color=COLORS["text"],
            bg_color="transparent",
            corner_radius=8,
            border_width=0,
        )
        body.pack()
        label = ctk.CTkLabel(
            body,
            text=self.text,
            justify="left",
            text_color="#ffffff",
            fg_color="transparent",
            wraplength=wraplength,
            font=label_font,
        )
        label.pack(padx=padx, pady=8)

        self.tooltip.update_idletasks()
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        width = self.tooltip.winfo_reqwidth()
        height = self.tooltip.winfo_reqheight()

        if x + width > screen_width - 8:
            x = max(8, screen_width - width - 8)
        if y + height > screen_height - 8:
            y = max(8, self.widget.winfo_rooty() - height - 8)

        self.tooltip.geometry(f"+{x}+{y}")
        self.tooltip.deiconify()
        self.tooltip.lift()
        HoverTooltip.active = self

    def _hide(self, _event: tk.Event | None = None) -> None:
        if _event is not None and self._pointer_inside_widget():
            return
        self._cancel()
        self._hide_now()

    def _hide_now(self) -> None:
        if self.tooltip is not None:
            try:
                self.tooltip.destroy()
            except (tk.TclError, AttributeError):
                pass
            self.tooltip = None
        if HoverTooltip.active is self:
            HoverTooltip.active = None

    def _font_family(self) -> str:
        if sys.platform.startswith("win"):
            return "Malgun Gothic"
        if sys.platform == "darwin":
            return "Apple SD Gothic Neo"
        return "TkDefaultFont"

    def _pointer_inside_widget(self) -> bool:
        x, y = self.widget.winfo_pointerxy()
        left = self.widget.winfo_rootx()
        top = self.widget.winfo_rooty()
        right = left + self.widget.winfo_width()
        bottom = top + self.widget.winfo_height()
        return left <= x <= right and top <= y <= bottom


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
        self.report_mode = tk.StringVar(value="direct")
        self.background = tk.BooleanVar(value=False)
        self.use_session = tk.BooleanVar(value=False)
        self.auto_retry = tk.BooleanVar(value=True)
        self.custom_filename = tk.BooleanVar(value=False)
        self.recommend_then_direct = tk.BooleanVar(value=False)
        self.filename_pattern = tk.StringVar(value=DEFAULT_FILENAME_PATTERN)
        self.filename_text_part = tk.StringVar(value="")
        self.filename_preview = tk.StringVar(value="")
        self.parallel_sessions = tk.StringVar(value=str(DEFAULT_PARALLEL_SESSIONS))
        self.status = tk.StringVar(value="대기 중")
        self.progress = tk.StringVar(value="0 / 0")
        self.progress_percent = tk.DoubleVar(value=0)
        self.success = tk.StringVar(value="0건")
        self.failed = tk.StringVar(value="0건")
        self.current_row = tk.StringVar(value="0")
        self.total_rows = tk.StringVar(value="0")

        self.start_button: ctk.CTkButton | None = None
        self.retry_failed_button: ctk.CTkButton | None = None
        self.stop_button: ctk.CTkButton | None = None
        self.force_stop_button: ctk.CTkButton | None = None
        self.status_badge: ctk.CTkLabel | None = None
        self.progress_bar: ctk.CTkProgressBar | None = None
        self.log_text: ctk.CTkTextbox | None = None
        self.background_switch: ctk.CTkSwitch | None = None
        self.use_session_switch: ctk.CTkSwitch | None = None
        self.auto_retry_switch: ctk.CTkSwitch | None = None
        self.filename_custom_switch: ctk.CTkSwitch | None = None
        self.filename_custom_frame: ctk.CTkFrame | None = None
        self.filename_parts_frame: ctk.CTkFrame | None = None
        self.filename_text_entry: ctk.CTkEntry | None = None
        self.filename_text_placeholder_label: ctk.CTkLabel | None = None
        self.filename_token_buttons: list[ctk.CTkButton] = []
        self.filename_text_buttons: list[ctk.CTkButton] = []
        self.filename_preset_buttons: list[ctk.CTkButton] = []
        self.filename_chip_widgets: list[ctk.CTkFrame] = []
        self.filename_chip_remove_buttons: list[ctk.CTkButton] = []
        self.filename_reset_button: ctk.CTkButton | None = None
        self.parallel_options_frame: ctk.CTkFrame | None = None
        self.parallel_sessions_menu: ctk.CTkOptionMenu | None = None
        self.report_mode_buttons: dict[str, ctk.CTkButton] = {}
        self.recommend_then_direct_frame: ctk.CTkFrame | None = None
        self.recommend_then_direct_switch: ctk.CTkSwitch | None = None
        self.tooltips: list[HoverTooltip] = []

        self.stop_requested = False
        self.force_stop_requested = False
        self.retry_failed_only_for_run = False
        self.auto_retry_runtime_enabled = True
        self.active_filename_pattern = ""
        self.filename_parts = [part.copy() for part in DEFAULT_FILENAME_PARTS]
        self.filename_drag_index: int | None = None
        self.filename_drag_start: tuple[int, int] | None = None
        self.filename_drag_active = False
        self.filename_drag_target_index: int | None = None
        self.filename_drag_ghost: tk.Toplevel | None = None
        self.parallel_options_enabled = False
        self.creator_click_count = 0
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_ui()
        self.filename_text_part.trace_add("write", lambda *_args: self._refresh_filename_text_placeholder())
        self._sync_filename_pattern_from_parts()
        self.after(200, self._poll_events)

    def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
        self.tooltips.append(HoverTooltip(widget, text))

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

        about_button = ctk.CTkButton(
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
        )
        about_button.pack(side="left")
        self._attach_tooltip(about_button, "프로그램 버전, 제작 정보, 문의 정보를 확인합니다.")

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
        input_button = ctk.CTkButton(
            panel,
            text="엑셀 선택",
            width=118,
            height=38,
            fg_color=COLORS["surface_alt"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._choose_input,
        )
        input_button.grid(row=0, column=2, sticky="e", padx=18, pady=(18, 8))
        self._attach_tooltip(input_button, "보고서 자동생성에 사용할 엑셀 파일을 선택합니다.")

        self._field_label(panel, "다운로드 위치").grid(row=1, column=0, sticky="w", padx=18, pady=(8, 8))
        ctk.CTkEntry(panel, textvariable=self.download_dir, height=38, border_color=COLORS["border"]).grid(
            row=1, column=1, sticky="ew", pady=(8, 8)
        )
        download_button = ctk.CTkButton(
            panel,
            text="폴더 선택",
            width=118,
            height=38,
            fg_color=COLORS["surface_alt"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._choose_download_dir,
        )
        download_button.grid(row=1, column=2, sticky="e", padx=18, pady=(8, 8))
        self._attach_tooltip(download_button, "완성된 PDF 보고서를 저장할 폴더를 선택합니다.")

        self.filename_custom_switch = ctk.CTkSwitch(
            panel,
            text="저장 파일명 커스텀",
            variable=self.custom_filename,
            command=self._on_filename_custom_toggled,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.filename_custom_switch.grid(row=2, column=0, columnspan=3, sticky="w", padx=18, pady=(6, 10))
        self._attach_tooltip(self.filename_custom_switch, "원하는 항목과 문자를 조합해 PDF 저장 파일명을 직접 지정합니다.")
        self._build_filename_custom_panel(panel)

    def _build_filename_custom_panel(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface_alt"],
            border_width=1,
            border_color="#e2e8f0",
            corner_radius=8,
        )
        frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=18, pady=(0, 18))
        frame.grid_columnconfigure(1, weight=1)

        self._field_label(frame, "파일명 구성").grid(row=0, column=0, sticky="nw", padx=14, pady=(14, 8))
        parts_row = ctk.CTkFrame(frame, fg_color="transparent")
        parts_row.grid(row=0, column=1, sticky="ew", padx=(8, 14), pady=(14, 8))
        parts_row.grid_columnconfigure(0, weight=1)
        self.filename_parts_frame = ctk.CTkFrame(
            parts_row,
            fg_color=COLORS["surface"],
            border_width=1,
            border_color=COLORS["border"],
            corner_radius=8,
        )
        self.filename_parts_frame.grid(row=0, column=0, sticky="ew")
        self.filename_parts_frame.configure(height=44)
        self.filename_parts_frame.grid_propagate(False)
        self.filename_parts_frame.bind("<Configure>", lambda _event: self._layout_filename_parts(), add="+")
        self.filename_reset_button = ctk.CTkButton(
            parts_row,
            text="초기화",
            width=72,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._reset_filename_parts,
        )
        self.filename_reset_button.grid(row=0, column=1, sticky="ne", padx=(8, 0))
        self._attach_tooltip(self.filename_reset_button, "파일명 구성을 모두 비우고 KOTRA 사이트 기본 파일명을 사용합니다.")

        ctk.CTkLabel(
            frame,
            text="추천 조합",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=1, column=0, sticky="nw", padx=14, pady=(6, 8))

        preset_frame = ctk.CTkFrame(frame, fg_color="transparent")
        preset_frame.grid(row=1, column=1, sticky="ew", padx=(8, 14), pady=(4, 8))
        self.filename_preset_buttons = []
        for label, parts in FILENAME_PRESETS:
            button = ctk.CTkButton(
                preset_frame,
                text=label,
                width=self._compact_button_width(label, min_width=86, extra=30),
                height=32,
                fg_color=COLORS["primary_soft"],
                hover_color="#dbeafe",
                border_width=1,
                border_color="#c7d8ff",
                text_color=COLORS["primary"],
                command=lambda item_parts=parts: self._apply_filename_preset(item_parts),
            )
            button.pack(side="left", padx=(0, 6), pady=3)
            self._attach_tooltip(button, "현재 파일명 구성을 이 추천 조합으로 바꿉니다.")
            self.filename_preset_buttons.append(button)

        ctk.CTkLabel(
            frame,
            text="항목 추가",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=2, column=0, sticky="nw", padx=14, pady=(6, 8))

        token_frame = ctk.CTkFrame(frame, fg_color="transparent")
        token_frame.grid(row=2, column=1, sticky="ew", padx=(8, 14), pady=(4, 8))

        self.filename_token_buttons = []
        token_rows = [
            ctk.CTkFrame(token_frame, fg_color="transparent"),
            ctk.CTkFrame(token_frame, fg_color="transparent"),
        ]
        for row_frame in token_rows:
            row_frame.pack(anchor="w", pady=(0, 4))

        for index, (label, token) in enumerate(FILENAME_PATTERN_TOKEN_LABELS):
            row_frame = token_rows[0 if index < 5 else 1]
            button = ctk.CTkButton(
                row_frame,
                text=label,
                width=self._compact_button_width(label),
                height=32,
                fg_color=COLORS["surface"],
                hover_color="#edf2f7",
                border_width=1,
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                command=lambda item_token=token: self._add_filename_token_part(item_token),
            )
            button.pack(side="left", padx=(0, 6), pady=3)
            self._attach_tooltip(button, f"파일명 구성에 {label} 항목을 추가합니다.")
            self.filename_token_buttons.append(button)

        ctk.CTkLabel(
            frame,
            text="문자 추가",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=3, column=0, sticky="nw", padx=14, pady=(6, 8))

        text_frame = ctk.CTkFrame(frame, fg_color="transparent")
        text_frame.grid(row=3, column=1, sticky="ew", padx=(8, 14), pady=(4, 8))
        text_frame.grid_columnconfigure(0, weight=1)
        self.filename_text_entry = ctk.CTkEntry(
            text_frame,
            textvariable=self.filename_text_part,
            height=32,
            border_color=COLORS["border"],
        )
        self.filename_text_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=3)
        self._attach_tooltip(self.filename_text_entry, "파일명에 넣을 고정 문구를 입력합니다.")
        self.filename_text_placeholder_label = ctk.CTkLabel(
            text_frame,
            text="추가할 문구를 입력하세요",
            text_color=COLORS["muted"],
            fg_color="transparent",
            font=ctk.CTkFont(size=13),
            height=18,
        )
        self.filename_text_placeholder_label.place(x=14, y=10)
        self.filename_text_placeholder_label.bind(
            "<Button-1>",
            lambda _event: self.filename_text_entry.focus_set() if self.filename_text_entry is not None else None,
            add="+",
        )

        add_text_button = ctk.CTkButton(
            text_frame,
            text="문자 추가",
            width=92,
            height=32,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._add_filename_custom_text_part,
        )
        add_text_button.grid(row=0, column=1, sticky="e", padx=(0, 0), pady=3)
        self._attach_tooltip(add_text_button, "입력한 고정 문구를 파일명 구성에 추가합니다.")
        self.filename_text_buttons = [add_text_button]

        shortcut_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        shortcut_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        for index, (label, value) in enumerate(FILENAME_TEXT_SHORTCUTS):
            button = ctk.CTkButton(
                shortcut_frame,
                text=label,
                width=58,
                height=28,
                fg_color=COLORS["surface"],
                hover_color="#edf2f7",
                border_width=1,
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                command=lambda item_value=value: self._add_filename_text_part(item_value),
            )
            button.grid(row=0, column=index, sticky="w", padx=(0 if index == 0 else 4, 0), pady=3)
            self._attach_tooltip(button, f"파일명 구성에 {label} 문자를 추가합니다.")
            self.filename_text_buttons.append(button)

        ctk.CTkLabel(
            frame,
            text="미리보기",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).grid(row=4, column=0, sticky="w", padx=14, pady=(4, 14))
        ctk.CTkLabel(
            frame,
            textvariable=self.filename_preview,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=13),
            anchor="w",
            justify="left",
            wraplength=760,
        ).grid(row=4, column=1, sticky="ew", padx=(8, 14), pady=(4, 14))

        self.filename_custom_frame = frame
        self._render_filename_parts()
        frame.grid_remove()

    def _build_options_panel(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 2, "실행 옵션")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_columnconfigure(1, weight=1)

        mode_box = ctk.CTkFrame(panel, fg_color="transparent")
        mode_box.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        mode_box.grid_columnconfigure(0, weight=1)
        self._field_label(mode_box, "보고서 생성 방식").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self._build_report_mode_selector(mode_box).grid(row=1, column=0, sticky="ew")
        recommend_option = self._build_recommend_then_direct_option(mode_box)
        recommend_option.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self._refresh_recommend_then_direct_visibility()

        switch_box = ctk.CTkFrame(
            panel,
            fg_color=COLORS["surface_alt"],
            border_width=1,
            border_color="#e2e8f0",
            corner_radius=8,
        )
        switch_box.grid(row=0, column=1, sticky="new", padx=18, pady=(18, 8))
        switch_box.grid_columnconfigure(0, weight=1)
        self.background_switch = ctk.CTkSwitch(
            switch_box,
            text="백그라운드 실행",
            variable=self.background,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        )
        self.background_switch.grid(row=0, column=0, sticky="w", padx=16, pady=(14, 10))
        self._attach_tooltip(self.background_switch, "브라우저 창을 띄우지 않고 백그라운드에서 실행합니다.")
        self.use_session_switch = ctk.CTkSwitch(
            switch_box,
            text="브라우저 세션 저장 사용",
            variable=self.use_session,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        )
        self.use_session_switch.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 10))
        self._attach_tooltip(self.use_session_switch, "저장된 로그인 세션을 사용하고, 실행 후 세션을 다시 저장합니다.")
        self.auto_retry_switch = ctk.CTkSwitch(
            switch_box,
            text="실패 항목 자동 재시도(1회)",
            variable=self.auto_retry,
            command=self._on_auto_retry_changed,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        )
        self.auto_retry_switch.grid(row=2, column=0, sticky="w", padx=16, pady=(0, 14))
        self.auto_retry_switch.select()
        self._attach_tooltip(self.auto_retry_switch, "실패한 항목을 자동으로 한 번 더 시도합니다. 실행 중 변경하면 이후 실패부터 반영됩니다.")

        self._build_parallel_options(panel)

        hint = ctk.CTkLabel(
            panel,
            text="자동 재시도는 실행 중 변경할 수 있으며 이후 새 실패부터 반영됩니다. 실패 행 다시 실행은 logs/failed_rows.xlsx 기준입니다.",
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
        self._attach_tooltip(self.parallel_sessions_menu, "개발자 옵션입니다. 동시에 실행할 브라우저 세션 수를 선택합니다.")

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
        self._attach_tooltip(self.start_button, "선택된 입력 파일과 옵션으로 보고서 자동 생성을 시작합니다.")

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
        self._attach_tooltip(self.stop_button, "현재 처리 중인 행을 마친 뒤 다음 행으로 넘어가지 않고 멈춥니다.")

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
        self._attach_tooltip(self.force_stop_button, "현재 작업을 가능한 즉시 중단합니다. 진행 중인 행은 실패로 남을 수 있습니다.")

        template_button = ctk.CTkButton(
            panel,
            text="입력 템플릿 생성",
            width=148,
            height=42,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._create_template,
        )
        template_button.grid(row=0, column=1, sticky="e")
        self._attach_tooltip(template_button, "입력 형식을 확인할 수 있는 예시 엑셀 템플릿을 생성합니다.")

        folders = ctk.CTkFrame(panel, fg_color="transparent")
        folders.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        open_download_button = ctk.CTkButton(
            folders,
            text="다운로드 폴더 열기",
            width=148,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._open_download_dir,
        )
        open_download_button.pack(side="left")
        self._attach_tooltip(open_download_button, "PDF 보고서가 저장되는 다운로드 폴더를 엽니다.")

        self.retry_failed_button = ctk.CTkButton(
            folders,
            text="실패 행 다시 실행",
            width=148,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._start_retry_failed,
        )
        self.retry_failed_button.pack(side="left", padx=(8, 0))
        self._attach_tooltip(self.retry_failed_button, "logs/failed_rows.xlsx에 남은 실패 항목만 다시 실행합니다.")

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
        self._bind_log_mousewheel()

        log_actions = ctk.CTkFrame(panel, fg_color="transparent")
        log_actions.grid(row=1, column=0, sticky="e", padx=18, pady=(0, 18))
        open_log_button = ctk.CTkButton(
            log_actions,
            text="로그 폴더 열기",
            width=128,
            height=36,
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=self._open_log_dir,
        )
        open_log_button.pack(side="right")
        self._attach_tooltip(open_log_button, "실행 기록, 실패 목록, 처리 상태 파일이 저장된 로그 폴더를 엽니다.")

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

    def _compact_button_width(self, text: str, min_width: int = 48, extra: int = 26) -> int:
        return max(min_width, min(150, len(text) * 12 + extra))

    def _build_report_mode_selector(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        selector = ctk.CTkFrame(parent, fg_color="#eef3f9", corner_radius=10)
        for column in range(2):
            selector.grid_columnconfigure(column, weight=1)
        self.report_mode_buttons = {}
        modes = [
            ("유망 시장 추천 보고서 생성", "recommend"),
            ("수출시장 분석 보고서 생성", "direct"),
        ]
        for column, (label, value) in enumerate(modes):
            button = ctk.CTkButton(
                selector,
                text=label,
                height=38,
                corner_radius=8,
                border_width=0,
                command=lambda selected=value: self._select_report_mode(selected),
                font=ctk.CTkFont(size=13, weight="bold"),
            )
            button.grid(row=0, column=column, sticky="ew", padx=(4 if column == 0 else 2, 4), pady=4)
            self.report_mode_buttons[value] = button
            if value == "direct":
                self._attach_tooltip(button, "희망 진출국을 기준으로 국가별 수출시장 분석 보고서를 생성합니다.")
            else:
                self._attach_tooltip(button, "유망 시장 추천 보고서를 생성합니다. 기본값은 추천 보고서만 저장합니다.")
        self._refresh_report_mode_buttons()
        return selector

    def _build_recommend_then_direct_option(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(
            parent,
            fg_color=COLORS["surface_alt"],
            border_width=1,
            border_color="#e2e8f0",
            corner_radius=8,
        )
        frame.grid_columnconfigure(0, weight=1)
        self.recommend_then_direct_switch = ctk.CTkSwitch(
            frame,
            text="추천 국가로 수출시장 분석보고서 생성 연동",
            variable=self.recommend_then_direct,
            fg_color="#cbd5e1",
            progress_color=COLORS["primary"],
            button_color="#ffffff",
            button_hover_color="#f8fafc",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=14),
        )
        self.recommend_then_direct_switch.grid(row=0, column=0, sticky="w", padx=14, pady=12)
        self._attach_tooltip(
            self.recommend_then_direct_switch,
            "추천 보고서 생성 후 추천 국가를 추출해 수출시장 분석보고서 생성까지 이어서 실행합니다.",
        )
        self.recommend_then_direct_frame = frame
        if self.report_mode.get() != "recommend":
            frame.grid_remove()
        return frame

    def _select_report_mode(self, value: str) -> None:
        self.report_mode.set(value)
        self._refresh_report_mode_buttons()
        self._refresh_recommend_then_direct_visibility()

    def _refresh_report_mode_buttons(self) -> None:
        for value, button in self.report_mode_buttons.items():
            selected = self.report_mode.get() == value
            button.configure(
                fg_color="#ffffff" if selected else "transparent",
                hover_color="#ffffff" if selected else "#e2eaf3",
                text_color=COLORS["primary"] if selected else COLORS["muted"],
            )

    def _refresh_recommend_then_direct_visibility(self) -> None:
        if self.recommend_then_direct_frame is None:
            return
        if self.report_mode.get() == "recommend":
            self.recommend_then_direct_frame.grid()
        else:
            self.recommend_then_direct_frame.grid_remove()

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

    def _bind_log_mousewheel(self) -> None:
        if self.log_text is None:
            return

        widgets = [self.log_text]
        for attr in ("_textbox", "_canvas", "_y_scrollbar"):
            widget = getattr(self.log_text, attr, None)
            if widget is not None:
                widgets.append(widget)

        for widget in widgets:
            widget.bind("<MouseWheel>", self._on_log_mousewheel, add="+")
            widget.bind("<Button-4>", self._on_log_mousewheel, add="+")
            widget.bind("<Button-5>", self._on_log_mousewheel, add="+")

    def _on_log_mousewheel(self, event: tk.Event) -> str:
        if self.log_text is None:
            return "break"

        if getattr(event, "num", None) == 4:
            step = -3
        elif getattr(event, "num", None) == 5:
            step = 3
        elif event.delta == 0:
            return "break"
        elif sys.platform == "darwin":
            step = -event.delta
        else:
            step = -int(event.delta / 120) * 3

        if step != 0:
            self.log_text.yview_scroll(step, "units")
        return "break"

    def _on_filename_custom_toggled(self) -> None:
        if self.filename_custom_frame is None:
            return
        if self.custom_filename.get():
            self.filename_custom_frame.grid()
        else:
            self.filename_custom_frame.grid_remove()
        self._sync_filename_pattern_from_parts()
        self._set_filename_custom_controls_state(running=bool(self.worker and self.worker.is_alive()))

    def _add_filename_token_part(self, token: str) -> None:
        self.filename_parts.append({"type": "token", "value": token})
        self._render_filename_parts()
        self._sync_filename_pattern_from_parts()

    def _add_filename_text_part(self, text: str) -> None:
        if not text:
            return
        self.filename_parts.append({"type": "text", "value": text})
        self._render_filename_parts()
        self._sync_filename_pattern_from_parts()

    def _add_filename_custom_text_part(self) -> None:
        self._add_filename_text_part(self.filename_text_part.get())

    def _apply_filename_preset(self, parts: list[dict[str, str]]) -> None:
        self.filename_parts = [part.copy() for part in parts]
        self._render_filename_parts()
        self._sync_filename_pattern_from_parts()

    def _refresh_filename_text_placeholder(self) -> None:
        if self.filename_text_placeholder_label is None:
            return
        if self.filename_text_part.get():
            self.filename_text_placeholder_label.place_forget()
        else:
            self.filename_text_placeholder_label.place(x=14, y=10)

    def _reset_filename_parts(self) -> None:
        self.filename_parts = []
        self._render_filename_parts()
        self._sync_filename_pattern_from_parts()

    def _remove_filename_part(self, index: int) -> None:
        if 0 <= index < len(self.filename_parts):
            self.filename_parts.pop(index)
            self._render_filename_parts()
            self._sync_filename_pattern_from_parts()

    def _filename_part_pattern(self, part: dict[str, str]) -> str:
        if part.get("type") == "token":
            return f"{{{part.get('value', '')}}}"
        return part.get("value", "")

    def _filename_part_label(self, part: dict[str, str]) -> str:
        if part.get("type") == "token":
            return FILENAME_TOKEN_LABEL_BY_TOKEN.get(part.get("value", ""), part.get("value", ""))
        value = part.get("value", "")
        if value == " ":
            return "공백"
        return value

    def _filename_pattern_from_parts(self) -> str:
        return "".join(self._filename_part_pattern(part) for part in self.filename_parts)

    def _sync_filename_pattern_from_parts(self) -> None:
        self.filename_pattern.set(self._filename_pattern_from_parts())
        self._update_filename_preview()

    def _render_filename_parts(self) -> None:
        if self.filename_parts_frame is None:
            return

        for child in self.filename_parts_frame.winfo_children():
            child.destroy()

        self.filename_chip_widgets = []
        self.filename_chip_remove_buttons = []
        if not self.filename_parts:
            empty_label = ctk.CTkLabel(
                self.filename_parts_frame,
                text="파일명 구성이 비어 있으면 사이트 기본 파일명으로 저장합니다.",
                text_color=COLORS["muted"],
                font=ctk.CTkFont(size=13),
            )
            empty_label.place(x=10, y=10)
            self.filename_parts_frame.configure(height=44)
            return

        for index, part in enumerate(self.filename_parts):
            chip = ctk.CTkFrame(
                self.filename_parts_frame,
                fg_color="#e8f0ff" if part.get("type") == "token" else "#ffffff",
                border_width=1,
                border_color=COLORS["border"],
                corner_radius=6,
            )
            chip.place(x=4, y=4)

            label = ctk.CTkLabel(
                chip,
                text=self._filename_part_label(part),
                text_color=COLORS["primary"] if part.get("type") == "token" else COLORS["text"],
                font=ctk.CTkFont(size=12, weight="bold" if part.get("type") == "token" else "normal"),
                anchor="w",
            )
            label.grid(row=0, column=0, sticky="w", padx=(8, 2), pady=4)

            remove_button = ctk.CTkButton(
                chip,
                text="×",
                width=22,
                height=20,
                fg_color="transparent",
                hover_color="#e2e8f0",
                text_color=COLORS["muted"],
                command=lambda item_index=index: self._remove_filename_part(item_index),
            )
            remove_button.grid(row=0, column=1, sticky="e", padx=(0, 4), pady=3)
            self._attach_tooltip(remove_button, "이 조각을 삭제합니다.")

            for widget in (chip, label):
                widget.bind("<ButtonPress-1>", lambda event, item_index=index: self._start_filename_part_drag(event, item_index), add="+")
                widget.bind("<B1-Motion>", self._move_filename_part_drag, add="+")
                widget.bind("<ButtonRelease-1>", self._finish_filename_part_drag, add="+")

            self.filename_chip_widgets.append(chip)
            self.filename_chip_remove_buttons.append(remove_button)

        self.after_idle(self._layout_filename_parts)
        self._set_filename_custom_controls_state(running=bool(self.worker and self.worker.is_alive()))

    def _layout_filename_parts(self) -> None:
        if self.filename_parts_frame is None or not self.filename_chip_widgets:
            return

        available_width = max(120, self.filename_parts_frame.winfo_width() - 8)
        x = 4
        y = 4
        row_height = 0
        gap = 6

        for chip in self.filename_chip_widgets:
            chip.update_idletasks()
            width = chip.winfo_reqwidth()
            height = chip.winfo_reqheight()
            if x > 4 and x + width > available_width:
                x = 4
                y += row_height + gap
                row_height = 0
            chip.place_configure(x=x, y=y)
            x += width + gap
            row_height = max(row_height, height)

        self.filename_parts_frame.configure(height=max(44, y + row_height + 4))

    def _start_filename_part_drag(self, event: tk.Event, index: int) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.filename_drag_index = index
        self.filename_drag_start = (event.x_root, event.y_root)
        self.filename_drag_active = False
        self.filename_drag_target_index = None

    def _move_filename_part_drag(self, event: tk.Event) -> None:
        if self.filename_drag_index is None or (self.worker and self.worker.is_alive()):
            return
        if not self.filename_drag_active:
            if not self._filename_drag_threshold_met(event.x_root, event.y_root):
                return
            self.filename_drag_active = True
            self._show_filename_drag_ghost(self.filename_drag_index)
            self._refresh_filename_drag_highlight()

        self._move_filename_drag_ghost(event.x_root, event.y_root)
        target_index = self._filename_part_index_at_pointer(event.x_root, event.y_root, self.filename_drag_index)
        if target_index != self.filename_drag_target_index:
            self.filename_drag_target_index = target_index
            self._refresh_filename_drag_highlight()

    def _finish_filename_part_drag(self, event: tk.Event) -> None:
        source_index = self.filename_drag_index
        was_dragging = self.filename_drag_active
        self.filename_drag_index = None
        self.filename_drag_start = None
        self.filename_drag_active = False
        self._hide_filename_drag_ghost()

        if source_index is None or (self.worker and self.worker.is_alive()):
            self.filename_drag_target_index = None
            self._refresh_filename_drag_highlight()
            return
        if not was_dragging:
            return

        target_index = self._filename_part_index_at_pointer(event.x_root, event.y_root, source_index)
        self.filename_drag_target_index = None
        if target_index is None or target_index == source_index:
            self._refresh_filename_drag_highlight()
            return

        part = self.filename_parts.pop(source_index)
        self.filename_parts.insert(max(0, target_index), part)
        self._render_filename_parts()
        self._sync_filename_pattern_from_parts()

    def _filename_drag_threshold_met(self, x_root: int, y_root: int) -> bool:
        if self.filename_drag_start is None:
            return False
        start_x, start_y = self.filename_drag_start
        return abs(x_root - start_x) >= 6 or abs(y_root - start_y) >= 6

    def _refresh_filename_drag_highlight(self) -> None:
        for index, widget in enumerate(self.filename_chip_widgets):
            if index >= len(self.filename_parts):
                continue
            part = self.filename_parts[index]
            is_drag_source = index == self.filename_drag_index and self.filename_drag_active
            is_drag_target = index == self.filename_drag_target_index and index != self.filename_drag_index
            widget.configure(
                fg_color="#f8fafc" if is_drag_source else ("#e8f0ff" if part.get("type") == "token" else "#ffffff"),
                border_color=COLORS["primary"] if is_drag_target else COLORS["border"],
            )

    def _show_filename_drag_ghost(self, index: int) -> None:
        self._hide_filename_drag_ghost()
        if not (0 <= index < len(self.filename_parts)):
            return

        part = self.filename_parts[index]
        ghost = tk.Toplevel(self)
        ghost.withdraw()
        ghost.overrideredirect(True)
        ghost.attributes("-topmost", True)
        if sys.platform == "darwin":
            ghost.configure(bg="systemTransparent")
            ghost.wm_attributes("-transparent", True)
        else:
            ghost.configure(bg=COLORS["surface"])

        chip = ctk.CTkFrame(
            ghost,
            fg_color="#e8f0ff" if part.get("type") == "token" else "#ffffff",
            border_width=1,
            border_color=COLORS["primary"],
            corner_radius=6,
        )
        chip.pack()
        ctk.CTkLabel(
            chip,
            text=self._filename_part_label(part),
            text_color=COLORS["primary"] if part.get("type") == "token" else COLORS["text"],
            font=ctk.CTkFont(size=12, weight="bold" if part.get("type") == "token" else "normal"),
        ).pack(padx=10, pady=5)

        self.filename_drag_ghost = ghost

    def _move_filename_drag_ghost(self, x_root: int, y_root: int) -> None:
        if self.filename_drag_ghost is None:
            return
        self.filename_drag_ghost.geometry(f"+{x_root + 12}+{y_root + 10}")
        self.filename_drag_ghost.deiconify()
        self.filename_drag_ghost.lift()

    def _hide_filename_drag_ghost(self) -> None:
        if self.filename_drag_ghost is None:
            return
        try:
            self.filename_drag_ghost.destroy()
        except tk.TclError:
            pass
        self.filename_drag_ghost = None

    def _filename_part_index_at_pointer(self, x_root: int, y_root: int, source_index: int) -> int | None:
        target_index: int | None = None
        nearest_distance: float | None = None
        for index, widget in enumerate(self.filename_chip_widgets):
            if index == source_index:
                continue
            left = widget.winfo_rootx()
            top = widget.winfo_rooty()
            width = max(1, widget.winfo_width())
            height = max(1, widget.winfo_height())
            right = left + width
            bottom = top + height
            if left <= x_root <= right and top <= y_root <= bottom:
                return index

            center_x = left + width / 2
            center_y = top + height / 2
            distance = ((x_root - center_x) ** 2) + ((y_root - center_y) ** 2)
            if nearest_distance is None or distance < nearest_distance:
                nearest_distance = distance
                target_index = index
        return target_index

    def _update_filename_preview(self) -> None:
        if not self.custom_filename.get():
            self.filename_preview.set("KOTRA 사이트 기본 파일명을 사용합니다.")
            return

        pattern = self._filename_pattern_from_parts()
        try:
            preview = render_filename_pattern(
                pattern,
                FILENAME_PREVIEW_ROW,
                suggested_filename="베트남_스킨케어(330499)_수출시장분석보고서.pdf",
            )
        except ValueError as exc:
            self.filename_preview.set(str(exc))
            return
        self.filename_preview.set(preview)

    def _filename_pattern_for_run(self) -> str | None:
        if not self.custom_filename.get():
            return ""

        pattern = self._filename_pattern_from_parts().strip()
        if not pattern:
            return ""

        try:
            render_filename_pattern(
                pattern,
                FILENAME_PREVIEW_ROW,
                suggested_filename="베트남_스킨케어(330499)_수출시장분석보고서.pdf",
            )
        except ValueError as exc:
            messagebox.showerror("파일명 패턴 오류", str(exc))
            return None
        return pattern

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
        opened_file, _ = self._open_path(path)
        if not opened_file:
            self._open_path(path.parent)

        messagebox.showinfo(
            "완료",
            f"입력 템플릿 저장 위치\n{path}",
        )

    def _open_download_dir(self) -> None:
        self._open_folder(Path(self.download_dir.get()))

    def _open_log_dir(self) -> None:
        self._open_folder(DEFAULT_LOG_DIR)

    def _open_folder(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        opened, exc = self._open_path(path)
        if not opened:
            messagebox.showerror("폴더 열기 실패", f"폴더를 열 수 없습니다.\n{path}\n\n{exc}")

    def _open_path(self, path: Path) -> tuple[bool, Exception | None]:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            return True, None
        except Exception as exc:
            return False, exc

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
        self._start_run(retry_failed_only=False)

    def _start_retry_failed(self) -> None:
        self._start_run(retry_failed_only=True)

    def _start_run(self, retry_failed_only: bool) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("실행 중", "이미 작업이 실행 중입니다.")
            return

        input_path = Path(self.input_path.get())
        if not retry_failed_only and not input_path.exists():
            messagebox.showerror("오류", f"엑셀 파일을 찾을 수 없습니다.\n{input_path}")
            return

        filename_pattern = self._filename_pattern_for_run()
        if filename_pattern is None:
            return
        self.active_filename_pattern = filename_pattern
        self.retry_failed_only_for_run = retry_failed_only

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
        start_message = "실패 행 다시 실행을 시작합니다." if retry_failed_only else "작업을 시작합니다."
        self._append_log(start_message, "info")
        self._sync_auto_retry_runtime(log_change=False)
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
        return self.retry_failed_only_for_run

    def _set_running_state(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state="disabled" if running else "normal")
        if self.retry_failed_button is not None:
            self.retry_failed_button.configure(state="disabled" if running else "normal")
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
        for button in self.report_mode_buttons.values():
            button.configure(state="disabled" if running else "normal")
        if self.recommend_then_direct_switch is not None:
            self.recommend_then_direct_switch.configure(state="disabled" if running else "normal")
        for option_switch in (self.background_switch, self.use_session_switch):
            if option_switch is not None:
                option_switch.configure(state="disabled" if running else "normal")
        self._set_filename_custom_controls_state(running)

    def _set_filename_custom_controls_state(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        if self.filename_custom_switch is not None:
            self.filename_custom_switch.configure(state=state)
        control_state = "normal" if self.custom_filename.get() and not running else "disabled"
        if self.filename_text_entry is not None:
            self.filename_text_entry.configure(state=control_state)
        for button in self.filename_token_buttons:
            button.configure(state=control_state)
        for button in self.filename_text_buttons:
            button.configure(state=control_state)
        for button in self.filename_preset_buttons:
            button.configure(state=control_state)
        for button in self.filename_chip_remove_buttons:
            button.configure(state=control_state)
        if self.filename_reset_button is not None:
            self.filename_reset_button.configure(state=control_state)

    def _on_auto_retry_changed(self) -> None:
        self._sync_auto_retry_runtime(log_change=bool(self.worker and self.worker.is_alive()))

    def _sync_auto_retry_runtime(self, log_change: bool = True) -> None:
        enabled = bool(self.auto_retry.get())
        self.auto_retry_runtime_enabled = enabled
        if log_change:
            state = "켜짐" if enabled else "꺼짐"
            self._append_log(f"자동 재시도 설정 변경: {state}", "warning")

    def _auto_retry_enabled(self) -> bool:
        return self.auto_retry_runtime_enabled

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
                row_retry_count=DEFAULT_ROW_RETRY_COUNT,
                auto_retry_enabled=self._auto_retry_enabled,
                filename_pattern=self.active_filename_pattern,
                report_mode=self.report_mode.get(),
                recommend_then_direct=self.report_mode.get() == "recommend" and self.recommend_then_direct.get(),
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
        stopped = bool(result.get("stopped", False))
        force_stopped = bool(result.get("force_stopped", False))
        if force_stopped:
            status_text = "강제종료됨"
            badge = "danger"
            log_text = "작업이 강제종료되었습니다. 처리 중이던 행은 실패로 기록되었습니다."
            log_tag = "danger"
            title = "강제종료됨"
        elif stopped:
            status_text = "중지됨"
            badge = "warning"
            log_text = "작업이 중지되었습니다."
            log_tag = "warning"
            title = "중지됨"
        else:
            status_text = "완료"
            badge = "warning" if failed else "success"
            log_text = "작업이 완료되었습니다."
            log_tag = "success" if failed == 0 else "warning"
            title = "완료"

        self.status.set(status_text)
        self._set_status_badge(badge)
        self._append_log(log_text, log_tag)
        self._set_running_state(False)
        messagebox.showinfo(
            title,
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
