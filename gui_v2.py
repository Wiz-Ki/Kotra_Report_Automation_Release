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

from automation import (
    FILENAME_PATTERN_TOKEN_LABELS,
    render_filename_pattern,
    run_automation,
    split_country_values,
)
from config import (
    APP_CREDITS,
    BASE_DIR,
    DEFAULT_DIRECT_REPORT_COUNT,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_LOG_DIR,
    DEFAULT_PARALLEL_SESSIONS,
    DEFAULT_ROW_RETRY_COUNT,
    DEFAULT_STATE_PATH,
    MAX_DIRECT_REPORT_COUNT,
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

_STATUS_BADGE_COLORS = {
    "처리 중": ("#dbeafe", COLORS["primary"], "#3b82f6"),
    "처리완료": ("#dcfce7", COLORS["success"], "#65a30d"),
    "처리실패": ("#fee2e2", COLORS["danger"], "#ef4444"),
    "자동 재시도 대기": ("#fef3c7", COLORS["warning"], "#d97706"),
    "처리 안됨": ("#e5e7eb", COLORS["muted"], "#9ca3af"),
}

_TABLE_COLUMN_LAYOUT = [
    (0, 44, 0),    # #
    (1, 220, 6),   # 상품명
    (2, 96, 2),    # HS코드
    (3, 120, 2),   # 대상국
    (4, 64, 1),    # 세션
    (5, 126, 2),   # 상태
    (6, 82, 1),    # 경과
    (7, 72, 1),    # 파일
]
_RESIZE_HANDLE_COLOR = "#d8e0eb"
_RESIZE_HANDLE_HOVER_COLOR = COLORS["primary"]
_RESIZE_HANDLE_BG_HOVER = "#f3f7ff"

# 자식(세부 작업) 행 관련 상수
# 펼침/접힘 애니메이션은 원격(VM) 환경에서 화면 갱신 부하를 키워 제거했다(즉시 표시/숨김).
_CHILD_ROW_HEIGHT = 30          # _make_child_row 의 height 와 일치해야 함
_CHILD_ROW_GAP = 1              # 자식 행 pack pady 하단 간격

# 진행 테이블을 한 번에 다 만들면 행이 많은 엑셀에서 메인 스레드가 수 초간 멈춰
# 윈도우에서 '응답 없음'이 뜬다. after 로 잘게 나눠 만들어 UI 응답성을 유지한다.
_TABLE_BUILD_CHUNK_ROWS = 12     # 한 틱에 생성할 행 수
_TABLE_BUILD_CHUNK_DELAY_MS = 8  # 청크 사이 양보 간격(ms)
_MAX_EVENTS_PER_POLL = 80        # 폴링 한 번에 처리할 워커 이벤트 상한(폭주 보호)
_LOG_MAX_LINES = 1500            # 로그 textbox 최대 줄 수(장시간 실행 시 느려짐 방지)

# 진행 테이블: 헤더는 스크롤 영역 밖이라 본문(CTkScrollableFrame) 내부보다
# 세로 스크롤바 폭만큼 넓다. 스크롤바가 보일 때만 헤더 우측에 이만큼 거터를 둬서
# 열 정렬을 맞춘다(스크롤바가 숨겨지면 본문이 넓어지므로 거터도 0이 되어야 함).
_SCROLLBAR_GUTTER = 16
_SCROLLBAR_CHECK_INTERVAL = 250  # 스크롤바 필요 여부 점검 주기(ms) — 로그창과 동일한 방식
# 데이터가 좌측 정렬인 열(상품명). 헤더 글자도 같은 정렬로 맞춰야 어긋나 보이지 않음.
# 대상국(3)은 HS코드처럼 헤더/데이터 모두 중앙 정렬이므로 포함하지 않는다.
_TABLE_LEFT_ALIGNED_COLUMNS = {1}

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
        self.direct_report_count = tk.StringVar(value=str(DEFAULT_DIRECT_REPORT_COUNT))
        self.filename_pattern = tk.StringVar(value=DEFAULT_FILENAME_PATTERN)
        self.filename_text_part = tk.StringVar(value="")
        self.filename_preview = tk.StringVar(value="")
        self.parallel_sessions = tk.StringVar(value=str(DEFAULT_PARALLEL_SESSIONS))
        self.status = tk.StringVar(value="대기 중")
        self.progress = tk.StringVar(value="0 / 0")
        self.progress_percent = tk.DoubleVar(value=0)
        self.total_count = tk.StringVar(value="0")
        self.running_count = tk.StringVar(value="0")
        self.completed_count = tk.StringVar(value="0")
        self.failed_count = tk.StringVar(value="0")

        self.start_button: ctk.CTkButton | None = None
        self.retry_failed_button: ctk.CTkButton | None = None
        self.stop_button: ctk.CTkButton | None = None
        self.force_stop_button: ctk.CTkButton | None = None
        self.status_badge: ctk.CTkLabel | None = None
        self.progress_bar: ctk.CTkProgressBar | None = None
        self.log_text: ctk.CTkTextbox | None = None
        self._board_tabs: dict[str, ctk.CTkButton] = {}
        self._board_tab_lines: dict[str, ctk.CTkFrame] = {}
        self._board_tab_frames: dict[str, ctk.CTkFrame] = {}
        self._board_content: ctk.CTkFrame | None = None
        self._progress_body: ctk.CTkScrollableFrame | None = None
        self._progress_header: ctk.CTkFrame | None = None
        self._header_gutter = 0  # 스크롤바가 보일 때만 _SCROLLBAR_GUTTER 로 바뀜
        self._progress_empty_label: ctk.CTkLabel | None = None
        self._progress_row_widgets: dict[str, dict] = {}
        self._row_start_times: dict[str, float] = {}
        self._running_keys: set[str] = set()
        self._row_status_by_key: dict[str, str] = {}
        self._child_keys_by_parent: dict[str, list[str]] = {}
        self._child_status_by_key: dict[str, str] = {}
        self._child_state_by_key: dict[str, dict] = {}
        self._child_expanded_parents: set[str] = set()
        # 청크 빌드 중 도착한 이벤트 보관용(행 위젯 생성 직후 적용)
        self._deferred_row_payloads: dict[str, dict] = {}
        self._deferred_task_payloads: dict[str, dict[str, dict]] = {}
        self._table_build_job: str | None = None
        self._table_build_specs: list[tuple[str, dict]] = []
        self._table_build_index = 0
        self._table_font_cache: dict[str, ctk.CTkFont] = {}
        self._table_column_widths = {column: minsize for column, minsize, _weight in _TABLE_COLUMN_LAYOUT}
        self._table_column_min_widths = {column: max(36, int(minsize * 0.55)) for column, minsize, _weight in _TABLE_COLUMN_LAYOUT}
        self._table_frames: list[ctk.CTkFrame] = []
        self._column_resize_state: dict[str, int] | None = None
        self._page_frame: ctk.CTkScrollableFrame | None = None
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
        self.direct_report_count_menu: ctk.CTkOptionMenu | None = None
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
        self.after(1000, self._tick_elapsed)
        self.after(_SCROLLBAR_CHECK_INTERVAL, self._update_progress_scrollbar)

    def _attach_tooltip(self, widget: tk.Widget, text: str) -> None:
        self.tooltips.append(HoverTooltip(widget, text))

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        page = ctk.CTkScrollableFrame(self, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew", padx=24, pady=22)
        page.grid_columnconfigure(0, weight=1)
        self._page_frame = page

        self._build_header(page)
        self._build_file_panel(page)
        self._build_options_panel(page)
        self._build_action_panel(page)
        self._build_execution_board(page)

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

    def _build_execution_board(self, parent: ctk.CTkFrame) -> None:
        panel = self._section(parent, 4, "실행 현황")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(4, weight=1)

        metrics = ctk.CTkFrame(panel, fg_color="transparent")
        metrics.grid(row=0, column=0, sticky="ew")
        for column in range(4):
            metrics.grid_columnconfigure(column, weight=1)

        self._metric_card(metrics, 0, "전체", self.total_count, COLORS["text"])
        self._metric_card(metrics, 1, "처리중", self.running_count, COLORS["primary"])
        self._metric_card(metrics, 2, "완료", self.completed_count, COLORS["success"])
        self._metric_card(metrics, 3, "실패", self.failed_count, COLORS["danger"])

        progress_row = ctk.CTkFrame(panel, fg_color="transparent")
        progress_row.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 14))
        progress_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            progress_row,
            text="전체 진행률",
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
            width=88,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        self.progress_bar = ctk.CTkProgressBar(
            progress_row,
            height=10,
            fg_color="#e2e8f0",
            progress_color=COLORS["primary"],
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=(10, 12))
        self.progress_bar.set(0)
        ctk.CTkLabel(
            progress_row,
            textvariable=self.progress,
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
            width=92,
            anchor="e",
        ).grid(row=0, column=2, sticky="e")

        status_box = ctk.CTkFrame(panel, fg_color=COLORS["surface_alt"], corner_radius=8)
        status_box.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 16))
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

        self._build_board_tabs(panel)
        self._show_board_tab("progress")

        log_actions = ctk.CTkFrame(panel, fg_color="transparent")
        log_actions.grid(row=5, column=0, sticky="e", padx=18, pady=(10, 18))
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

    def _build_board_tabs(self, parent: ctk.CTkFrame) -> None:
        tab_bar = ctk.CTkFrame(parent, fg_color="transparent")
        tab_bar.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 0))

        self._board_tabs = {}
        self._board_tab_lines = {}
        for index, (key, label) in enumerate((("progress", "진행 현황"), ("log", "로그"))):
            tab_item = ctk.CTkFrame(tab_bar, fg_color="transparent")
            tab_item.grid(row=0, column=index, sticky="w", padx=(0 if index == 0 else 4, 0))
            button = ctk.CTkButton(
                tab_item,
                text=label,
                width=96,
                height=34,
                corner_radius=0,
                border_width=0,
                fg_color="transparent",
                hover_color=COLORS["surface_alt"],
                text_color=COLORS["muted"],
                font=ctk.CTkFont(size=14, weight="bold"),
                command=lambda selected=key: self._show_board_tab(selected),
            )
            button.grid(row=0, column=0, sticky="ew")
            line = ctk.CTkFrame(tab_item, fg_color="transparent", height=3)
            line.grid(row=1, column=0, sticky="ew", padx=8)
            line.grid_propagate(False)
            self._board_tabs[key] = button
            self._board_tab_lines[key] = line

        content = ctk.CTkFrame(parent, fg_color=COLORS["surface"], border_width=1, border_color=COLORS["border"], corner_radius=8)
        content.grid(row=4, column=0, sticky="nsew", padx=18, pady=(0, 0))
        content.configure(height=330)
        content.grid_propagate(False)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self._board_content = content

        progress_tab = ctk.CTkFrame(content, fg_color=COLORS["surface"])
        progress_tab.grid_columnconfigure(0, weight=1)
        progress_tab.grid_rowconfigure(1, weight=1)
        self._build_progress_tree(progress_tab)

        log_tab = ctk.CTkFrame(content, fg_color=COLORS["surface"])
        log_tab.grid_columnconfigure(0, weight=1)
        log_tab.grid_rowconfigure(0, weight=1)
        self.log_text = ctk.CTkTextbox(
            log_tab,
            height=220,
            fg_color="#0f172a",
            text_color="#dbeafe",
            border_width=0,
            corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=13),
            wrap="word",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log_text.configure(state="disabled")
        self._configure_log_tags()
        self._bind_log_mousewheel()

        self._board_tab_frames = {"progress": progress_tab, "log": log_tab}

    def _show_board_tab(self, selected: str) -> None:
        for key, frame in self._board_tab_frames.items():
            if key == selected:
                frame.grid(row=0, column=0, sticky="nsew")
            else:
                frame.grid_forget()

        for key, button in self._board_tabs.items():
            is_selected = key == selected
            button.configure(
                fg_color="transparent",
                hover_color=COLORS["surface_alt"],
                text_color=COLORS["text"] if is_selected else COLORS["muted"],
                border_width=0 if is_selected else 0,
            )
            line = self._board_tab_lines.get(key)
            if line is not None:
                line.configure(fg_color=COLORS["text"] if is_selected else "transparent")

    def _configure_table_cols(self, frame: ctk.CTkFrame) -> None:
        # 새 프레임에는 그 프레임의 열 설정만 적용한다. 전체 재적용(_apply_table_column_layout)을
        # 여기서 부르면 행 생성마다 모든 프레임을 다시 도는 O(n^2) Tcl 호출 폭주가 일어나
        # 행이 많은 엑셀에서 메인 스레드가 수 초간 멈춘다(윈도우 '응답 없음'의 원인).
        if frame not in self._table_frames:
            self._table_frames.append(frame)
        self._apply_columns_to_frame(frame)

    def _apply_columns_to_frame(self, frame: ctk.CTkFrame) -> None:
        gutter_col = len(_TABLE_COLUMN_LAYOUT)  # 실제 열(0~7) 뒤의 거터 열
        for column, _minsize, weight in _TABLE_COLUMN_LAYOUT:
            frame.grid_columnconfigure(
                column,
                minsize=self._table_column_widths.get(column, _minsize),
                weight=weight,
            )
        # 헤더는 스크롤 영역 밖이라, 스크롤바가 보일 때 본문 내부보다 그 폭만큼 넓어진다.
        # 그때만 헤더 우측에 거터를 둬서 가중치가 분배되는 실제 열 영역을 본문과
        # 동일하게 맞춘다(스크롤바가 숨겨지면 _header_gutter 가 0 이라 거터도 사라짐).
        is_header = frame is self._progress_header
        frame.grid_columnconfigure(
            gutter_col,
            minsize=(self._header_gutter if is_header else 0),
            weight=0,
        )

    def _apply_table_column_layout(self) -> None:
        """열 폭/거터가 바뀌었을 때만 모든 테이블 프레임에 일괄 재적용한다(리사이즈 등)."""
        live_frames = []
        for frame in self._table_frames:
            try:
                if not frame.winfo_exists():
                    continue
                self._apply_columns_to_frame(frame)
                live_frames.append(frame)
            except tk.TclError:
                continue
        self._table_frames = live_frames

    def _set_header_gutter(self, gutter: int) -> None:
        """스크롤바 표시 여부가 바뀔 때만 헤더 거터를 갱신하고 열 레이아웃을 다시 적용한다."""
        if self._header_gutter == gutter:
            return
        self._header_gutter = gutter
        self._apply_table_column_layout()

    def _update_progress_scrollbar(self) -> None:
        """진행 테이블 스크롤바를 로그창처럼 '필요할 때만' 표시한다.

        CTkScrollableFrame 은 스크롤바를 항상 띄우므로, CTkTextbox 와 같은 방식으로
        주기적으로 내용이 넘치는지(yview != 0~1) 확인해 grid / grid_forget 한다.
        스크롤바가 사라지면 본문이 그 폭만큼 넓어지므로 헤더 거터도 함께 맞춘다.
        """
        body = self._progress_body
        if body is not None:
            canvas = getattr(body, "_parent_canvas", None)
            scrollbar = getattr(body, "_scrollbar", None)
            if canvas is not None and scrollbar is not None:
                try:
                    # yview 는 내용이 보이는 영역보다 작은 상태에서 시점이 어긋나면
                    # 신뢰할 수 없으므로, 실제 높이 비교로 필요 여부를 판정한다.
                    needed = self._progress_scroll_needed()
                    mapped = bool(scrollbar.winfo_ismapped())
                    if needed and not mapped:
                        scrollbar.grid(row=1, column=1, sticky="nsew", pady=0)
                        self._set_header_gutter(_SCROLLBAR_GUTTER)
                    elif not needed and mapped:
                        scrollbar.grid_forget()
                        self._set_header_gutter(0)
                    # 내용이 다 보이는데 시점이 어긋나 있으면(경계 밖으로 끌린 상태) 원위치.
                    # 이 상태에서 yview() 는 (0.0, 1.0) 을 반환하므로(신뢰 불가),
                    # 실제 뷰 상단의 캔버스 좌표(canvasy)를 봐야 드리프트를 감지할 수 있다.
                    if not needed and abs(canvas.canvasy(0)) > 0.5:
                        canvas.yview_moveto(0)
                except tk.TclError:
                    pass
        self.after(_SCROLLBAR_CHECK_INTERVAL, self._update_progress_scrollbar)

    def _build_table_header_cell(self, parent: ctk.CTkFrame, column: int, text: str) -> None:
        cell = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        cell.grid(row=0, column=column, sticky="nsew", padx=(4 if column == 0 else 0, 0), pady=0)
        cell.grid_columnconfigure(0, weight=1)
        cell.grid_rowconfigure(0, weight=1)

        left_aligned = column in _TABLE_LEFT_ALIGNED_COLUMNS
        ctk.CTkLabel(
            cell,
            text=text,
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w" if left_aligned else "center",
            # 우측 16px 는 리사이즈 핸들 자리. 중앙 정렬 열은 좌우 패딩을 대칭(16,16)으로
            # 둬야 헤더 글자 중심이 데이터 글자 중심(=컬럼 중앙)과 정확히 일치한다.
        ).grid(row=0, column=0, sticky="nsew", padx=((4, 16) if left_aligned else (16, 16)), pady=4)

        if column >= len(_TABLE_COLUMN_LAYOUT) - 1:
            return

        handle = tk.Frame(
            cell,
            cursor="sb_h_double_arrow",
            bg=COLORS["surface_alt"],
            bd=0,
            highlightthickness=0,
        )
        handle.place(relx=1.0, rely=0, relheight=1.0, width=16, anchor="ne")
        grip = tk.Frame(
            handle,
            width=1,
            bg=_RESIZE_HANDLE_COLOR,
            bd=0,
            highlightthickness=0,
            cursor="sb_h_double_arrow",
        )
        grip.place(relx=0.5, rely=0.22, relheight=0.56, anchor="n")

        def highlight(_event: tk.Event | None = None) -> None:
            handle.configure(bg=_RESIZE_HANDLE_BG_HOVER)
            grip.configure(bg=_RESIZE_HANDLE_HOVER_COLOR, width=2)

        def unhighlight(_event: tk.Event | None = None) -> None:
            if self._column_resize_state is None:
                handle.configure(bg=COLORS["surface_alt"])
                grip.configure(bg=_RESIZE_HANDLE_COLOR, width=1)

        def end_resize(event: tk.Event) -> str:
            result = self._end_column_resize(event)
            unhighlight()
            return result

        for widget in (handle, grip):
            widget.bind("<Enter>", highlight)
            widget.bind("<Leave>", unhighlight)
            widget.bind("<ButtonPress-1>", lambda event, col=column: self._start_column_resize(event, col))
            widget.bind("<B1-Motion>", self._drag_column_resize)
            widget.bind("<ButtonRelease-1>", end_resize)
        self._attach_tooltip(handle, "드래그해서 컬럼 폭을 조정합니다.")

    def _start_column_resize(self, event: tk.Event, column: int) -> str:
        next_column = column + 1
        self._column_resize_state = {
            "column": column,
            "next_column": next_column,
            "x_root": int(event.x_root),
            "start_width": int(self._table_column_widths.get(column, 0)),
            "next_start_width": int(self._table_column_widths.get(next_column, 0)),
        }
        return "break"

    def _drag_column_resize(self, event: tk.Event) -> str:
        state = self._column_resize_state
        if state is None:
            return "break"

        column = state["column"]
        next_column = state["next_column"]
        start_width = state["start_width"]
        next_start_width = state["next_start_width"]
        dx = int(event.x_root) - state["x_root"]

        min_width = self._table_column_min_widths.get(column, 36)
        next_min_width = self._table_column_min_widths.get(next_column, 36)
        applied_dx = max(min_width - start_width, min(dx, next_start_width - next_min_width))

        self._table_column_widths[column] = start_width + applied_dx
        self._table_column_widths[next_column] = next_start_width - applied_dx
        self._apply_table_column_layout()
        return "break"

    def _end_column_resize(self, _event: tk.Event) -> str:
        self._column_resize_state = None
        return "break"

    def _build_progress_tree(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(parent, fg_color=COLORS["surface_alt"], corner_radius=0, height=36)
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        header.grid_propagate(False)
        self._progress_header = header  # _apply_table_column_layout 에서 헤더만 거터 처리
        self._configure_table_cols(header)
        for col, text in [
            (0, "#"),
            (1, "상품명"),
            (2, "HS코드"),
            (3, "대상국"),
            (4, "세션"),
            (5, "상태"),
            (6, "경과"),
            (7, "파일"),
        ]:
            self._build_table_header_cell(header, col, text)

        body = ctk.CTkScrollableFrame(
            parent, fg_color=COLORS["surface_alt"], corner_radius=0, border_width=0,
        )
        body.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)
        self._progress_body = body
        self._setup_progress_scroll_isolation(body)
        self._show_progress_empty_state()

    def _progress_scroll_needed(self) -> bool:
        """진행 테이블 내용이 보이는 영역보다 커서 실제로 스크롤이 필요한지."""
        body = self._progress_body
        canvas = getattr(body, "_parent_canvas", None)
        if body is None or canvas is None:
            return False
        try:
            return body.winfo_reqheight() > canvas.winfo_height() + 1
        except tk.TclError:
            return False

    def _scroll_progress_canvas(self, event: tk.Event) -> str:
        body_canvas = getattr(self._progress_body, "_parent_canvas", None)
        if body_canvas is None:
            return "break"
        # 내용이 모두 보이면 스크롤하지 않는다. scrollregion 이 캔버스보다 작은 상태에서
        # yview_scroll 을 호출하면 경계 클램프가 되지 않아 행들이 위/아래로 끌려가
        # 테이블 상단에 빈 공간이 생기는 버그가 있다.
        if not self._progress_scroll_needed():
            return "break"
        try:
            if getattr(event, "num", None) == 4:
                body_canvas.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                body_canvas.yview_scroll(3, "units")
            elif not event.delta:
                return "break"
            elif sys.platform == "darwin":
                body_canvas.yview_scroll(-int(event.delta), "units")
            else:
                body_canvas.yview_scroll(-int(event.delta / 120) * 3, "units")
        except tk.TclError:
            pass
        return "break"

    def _setup_progress_scroll_isolation(self, body: ctk.CTkScrollableFrame) -> None:
        """진행 테이블 내부 스크롤만 이 영역에서 처리한다."""
        body_canvas = getattr(body, "_parent_canvas", None)
        if body_canvas is None:
            return
        for widget in (body, body_canvas):
            widget.bind("<MouseWheel>", self._scroll_progress_canvas, add="+")
            widget.bind("<Button-4>", self._scroll_progress_canvas, add="+")
            widget.bind("<Button-5>", self._scroll_progress_canvas, add="+")

    # 행 위젯에 쓰는 폰트는 공유 캐시를 사용한다. CTkFont 를 위젯마다 새로 만들면
    # 생성/등록 비용이 쌓여(행 수백 개 × 라벨 수) 테이블 빌드가 눈에 띄게 느려진다.
    _TABLE_FONT_SPECS = {
        "body12": dict(size=12),
        "body13": dict(size=13),
        "mode10": dict(size=10),
        "toggle10b": dict(size=10, weight="bold"),
        "mono12": dict(family="Consolas", size=12),
        "glyph11": dict(size=11),
        "pill12b": dict(size=12, weight="bold"),
    }

    def _table_font(self, key: str) -> ctk.CTkFont:
        font = self._table_font_cache.get(key)
        if font is None:
            font = ctk.CTkFont(**self._TABLE_FONT_SPECS[key])
            self._table_font_cache[key] = font
        return font

    def _make_data_row(self, parent: ctk.CTkScrollableFrame, row_index: int,
                       product: str, hs: str, country: str, mode_lbl: str) -> dict:
        frame = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=4, height=38)
        frame.pack(fill="x", expand=False, pady=(0, 1), padx=0)
        frame.pack_propagate(False)
        self._configure_table_cols(frame)

        ctk.CTkLabel(frame, text=str(row_index), text_color=COLORS["muted"],
                     font=self._table_font("body12"), anchor="center",
                     ).grid(row=0, column=0, sticky="ew", padx=(8, 4))
        product_box = ctk.CTkFrame(frame, fg_color="transparent")
        product_box.grid(row=0, column=1, sticky="ew", padx=(4, 4))
        product_box.grid_columnconfigure(0, weight=1)
        product_box.grid_columnconfigure(1, weight=0)
        ctk.CTkLabel(product_box, text=product, text_color=COLORS["text"],
                     font=self._table_font("body13"), anchor="w",
                     ).grid(row=0, column=0, sticky="ew", columnspan=2)
        ctk.CTkLabel(product_box, text=mode_lbl, text_color=COLORS["muted"],
                     font=self._table_font("mode10"), anchor="w",
                     ).grid(row=1, column=0, sticky="ew", pady=(0, 1))
        child_toggle = ctk.CTkButton(
            product_box,
            text="",
            width=0,
            height=20,
            fg_color="transparent",
            hover_color=COLORS["surface_alt"],
            text_color=COLORS["muted"],
            font=self._table_font("toggle10b"),
            command=lambda: None,
        )
        ctk.CTkLabel(frame, text=hs, text_color=COLORS["muted"],
                     font=self._table_font("mono12"), anchor="center",
                     ).grid(row=0, column=2, sticky="ew", padx=(4, 4))
        ctk.CTkLabel(frame, text=country, text_color=COLORS["muted"],
                     font=self._table_font("body12"), anchor="center",
                     ).grid(row=0, column=3, sticky="ew", padx=(4, 4))
        session_lbl = ctk.CTkLabel(frame, text="—", text_color=COLORS["muted"],
                                   font=self._table_font("body12"), anchor="center")
        session_lbl.grid(row=0, column=4, sticky="ew", padx=(4, 4))
        status_lbl = self._status_pill(frame, "처리 안됨")
        status_lbl.grid(row=0, column=5, sticky="", padx=(4, 4))
        elapsed_lbl = ctk.CTkLabel(frame, text="—", text_color=COLORS["muted"],
                                   font=self._table_font("body12"), anchor="center")
        elapsed_lbl.grid(row=0, column=6, sticky="ew", padx=(4, 4))
        file_button = ctk.CTkButton(
            frame,
            text="—",
            font=self._table_font("body12"),
            width=46,
            height=26,
            fg_color="transparent",
            hover_color=COLORS["surface_alt"],
            border_width=0,
            text_color=COLORS["muted"],
            state="disabled",
            command=lambda: None,
        )
        file_button.grid(row=0, column=7, sticky="", padx=(4, 8))

        children_frame = ctk.CTkFrame(parent, fg_color=COLORS["surface_alt"], corner_radius=0)
        children_frame.grid_columnconfigure(0, weight=1)
        self._bind_progress_row_scroll(frame)

        return {
            "frame": frame,
            "session_lbl": session_lbl,
            "status_lbl": status_lbl,
            "elapsed_lbl": elapsed_lbl,
            "file_button": file_button,
            "children_frame": children_frame,
            "children_visible": False,
            "child_toggle": child_toggle,
        }

    def _child_label_for_key(self, child_key: str) -> str:
        # child_key 형식: f"{parent_key}:{task_type}:{country}" (예: "row:1:direct:베트남")
        parts = child_key.split(":")
        task_type = parts[-2] if len(parts) >= 2 else ""
        country = parts[-1] if parts else ""
        return "추천 보고서" if task_type == "recommend" else f"투자 보고서 - {country}"

    def _ensure_child_widgets(self, parent_widgets: dict, child_key: str) -> dict | None:
        """자식 행 위젯을 처음 보여질 때 만들고, 보관해 둔 표시 상태를 적용한다."""
        widgets = self._progress_row_widgets.get(child_key)
        if widgets is not None:
            return widgets
        try:
            widgets = self._make_child_row(
                parent_widgets["children_frame"], self._child_label_for_key(child_key)
            )
        except tk.TclError:
            return None
        self._progress_row_widgets[child_key] = widgets
        state = self._child_state_by_key.get(child_key)
        if state:
            try:
                self._configure_status_pill(widgets["status_lbl"], str(state.get("status", "처리 안됨")))
                widgets["elapsed_lbl"].configure(text=str(state.get("elapsed_str") or "—"))
                self._update_file_button(widgets, [str(p) for p in state.get("saved_files", [])])
            except Exception:
                pass
        return widgets

    def _make_child_row(self, parent: ctk.CTkFrame, task_label: str) -> dict:
        frame = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=4, height=_CHILD_ROW_HEIGHT)
        frame.pack_propagate(False)
        self._configure_table_cols(frame)

        ctk.CTkLabel(frame, text="↳", text_color=COLORS["muted"],
                     font=self._table_font("glyph11"), anchor="e", width=20,
                     ).grid(row=0, column=0, sticky="e", padx=(0, 6))
        ctk.CTkLabel(frame, text=task_label, text_color=COLORS["muted"],
                     font=self._table_font("body12"), anchor="w",
                     ).grid(row=0, column=1, sticky="ew", padx=(10, 4))
        status_lbl = self._status_pill(frame, "처리 안됨", width=108)
        status_lbl.grid(row=0, column=5, padx=(4, 4))
        elapsed_lbl = ctk.CTkLabel(frame, text="—", text_color=COLORS["muted"],
                                   font=self._table_font("body12"), anchor="center")
        elapsed_lbl.grid(row=0, column=6, sticky="ew", padx=(4, 4))
        file_button = ctk.CTkButton(
            frame,
            text="—",
            font=self._table_font("body12"),
            width=46,
            height=24,
            fg_color="transparent",
            hover_color=COLORS["surface_alt"],
            border_width=0,
            text_color=COLORS["muted"],
            state="disabled",
            command=lambda: None,
        )
        file_button.grid(row=0, column=7, sticky="", padx=(4, 8))
        self._bind_progress_row_scroll(frame)

        return {"frame": frame, "status_lbl": status_lbl, "elapsed_lbl": elapsed_lbl, "file_button": file_button}

    def _status_pill(self, parent: ctk.CTkFrame, status: str, width: int = 104) -> ctk.CTkLabel:
        fg, text, _dot = _STATUS_BADGE_COLORS.get(status, _STATUS_BADGE_COLORS["처리 안됨"])
        return ctk.CTkLabel(
            parent,
            text=f"● {status}",
            width=width,
            height=24,
            fg_color=fg,
            text_color=text,
            corner_radius=12,
            font=self._table_font("pill12b"),
            anchor="center",
        )

    def _configure_status_pill(self, label: ctk.CTkLabel, status: str) -> None:
        fg, text, _dot = _STATUS_BADGE_COLORS.get(status, _STATUS_BADGE_COLORS["처리 안됨"])
        label.configure(text=f"● {status}", fg_color=fg, text_color=text)

    def _show_progress_empty_state(self) -> None:
        body = self._progress_body
        if body is None:
            return
        for widget in body.winfo_children():
            widget.destroy()
        self._progress_empty_label = ctk.CTkLabel(
            body,
            text="아직 실행된 항목이 없습니다.",
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=13, weight="bold"),
            height=160,
        )
        self._progress_empty_label.pack(fill="x", expand=True, pady=34)

    def _bind_progress_row_scroll(self, frame: ctk.CTkFrame) -> None:
        if getattr(self._progress_body, "_parent_canvas", None) is None:
            return

        def bind_recursive(widget: tk.Widget) -> None:
            widget.bind("<MouseWheel>", self._scroll_progress_canvas, add="+")
            widget.bind("<Button-4>", self._scroll_progress_canvas, add="+")
            widget.bind("<Button-5>", self._scroll_progress_canvas, add="+")
            for child in widget.winfo_children():
                bind_recursive(child)

        bind_recursive(frame)

    def _update_summary_counts(self) -> None:
        running = sum(1 for status in self._row_status_by_key.values() if status == "처리 중")
        self.running_count.set(str(running))

    def _task_sort_key(self, child_key: str) -> int:
        # 추천 보고서만 맨 위로 올리고, 나머지(투자 보고서)는 동일 그룹(1)이라
        # sorted 의 안정 정렬 특성상 _child_keys_by_parent 에 등록된 순서
        # (= 실제 처리 순서)를 그대로 유지한다.
        task_type = child_key.split(":")[-2] if ":" in child_key else ""
        return 0 if task_type == "recommend" else 1

    def _toggle_child_tasks(self, parent_key: str) -> None:
        if parent_key in self._child_expanded_parents:
            self._child_expanded_parents.remove(parent_key)
        else:
            self._child_expanded_parents.add(parent_key)
        self._refresh_child_visibility(parent_key)

    def _refresh_child_visibility(self, parent_key: str) -> None:
        parent_widgets = self._progress_row_widgets.get(parent_key)
        if parent_widgets is None:
            return
        child_keys = sorted(self._child_keys_by_parent.get(parent_key, []), key=self._task_sort_key)
        toggle = parent_widgets.get("child_toggle")
        if len(child_keys) <= 1:
            self._child_expanded_parents.discard(parent_key)
            if toggle is not None:
                toggle.grid_forget()
            for key in child_keys:
                widgets = self._progress_row_widgets.get(key)
                if widgets is not None:
                    widgets["frame"].pack_forget()
            children_frame = parent_widgets["children_frame"]
            if parent_widgets.get("children_visible"):
                children_frame.pack_forget()
                parent_widgets["children_visible"] = False
            return

        # 자동 펼침 없음: 사용자가 토글을 눌렀을 때만 자식 행을 보여준다.
        # (원격/VM 환경에서 잦은 자동 펼침·접힘이 화면 갱신 부하를 키우는 것을 방지)
        expanded = parent_key in self._child_expanded_parents
        visible_keys = child_keys if expanded else []

        if toggle is not None:
            toggle.configure(
                text=("▼ 접기" if expanded else f"▶ 작업 {len(child_keys)}개"),
                command=lambda key=parent_key: self._toggle_child_tasks(key),
                width=76 if expanded else 92,
            )
            toggle.grid(row=1, column=1, sticky="e", padx=(8, 0), pady=(0, 1))

        # 보여줄 자식 행들을 컨테이너 안에 배치(또는 제거)한다.
        for key in child_keys:
            if key in visible_keys:
                continue
            widgets = self._progress_row_widgets.get(key)
            if widgets is not None:
                widgets["frame"].pack_forget()
        # 자식 위젯은 처음 보여질 때 만든다(지연 생성). 생성 순서가 정렬 순서와 다를 수
        # 있으므로 after= 로 매번 정렬 순서대로 재배치한다(이미 그 위치면 사실상 no-op).
        prev_frame = None
        for key in child_keys:
            if key not in visible_keys:
                continue
            widgets = self._ensure_child_widgets(parent_widgets, key)
            if widgets is None:
                continue
            frame = widgets["frame"]
            # 자식 행도 부모와 동일하게 풀 너비로 둔다. 들여쓰기(padx)를 주면
            # 프레임이 좁아져 상태/경과/파일 열이 부모와 어긋나므로,
            # 계층 표시는 ↳ 글리프와 라벨 들여쓰기로만 표현한다.
            pack_kwargs: dict = {"fill": "x", "expand": False, "pady": (0, _CHILD_ROW_GAP), "padx": 0}
            if prev_frame is not None:
                pack_kwargs["after"] = prev_frame
            frame.pack(**pack_kwargs)
            prev_frame = frame

        children_frame = parent_widgets["children_frame"]
        if visible_keys:
            if not parent_widgets.get("children_visible"):
                # 자식 컨테이너는 항상 자기 부모 행 바로 뒤에 온다.
                children_frame.pack(fill="x", expand=False, after=parent_widgets["frame"])
                parent_widgets["children_visible"] = True
        elif parent_widgets.get("children_visible"):
            children_frame.pack_forget()
            parent_widgets["children_visible"] = False

    def _update_file_button(self, widgets: dict, saved_files: list[str]) -> None:
        button = widgets.get("file_button")
        if button is None:
            return
        paths = [Path(path) for path in saved_files if str(path).strip()]
        if not paths:
            button.configure(
                text="—",
                state="disabled",
                fg_color="transparent",
                border_width=0,
                text_color=COLORS["muted"],
                command=lambda: None,
            )
            return

        target = paths[0] if len(paths) == 1 else paths[0].parent
        button.configure(
            text="열기",
            state="normal",
            fg_color=COLORS["surface"],
            hover_color="#edf2f7",
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            command=lambda path=target: self._open_path(path),
        )

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
            command=self._refresh_recommend_then_direct_visibility,
        )
        self.recommend_then_direct_switch.grid(row=0, column=0, sticky="w", padx=14, pady=12)
        self._attach_tooltip(
            self.recommend_then_direct_switch,
            "추천 보고서 생성 후 추천 국가를 추출해 수출시장 분석보고서 생성까지 이어서 실행합니다.",
        )
        count_box = ctk.CTkFrame(frame, fg_color="transparent")
        count_box.grid(row=0, column=1, sticky="e", padx=14, pady=10)
        ctk.CTkLabel(
            count_box,
            text="분석보고서 수",
            text_color=COLORS["muted"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=(0, 8))

        self.direct_report_count_menu = ctk.CTkOptionMenu(
            count_box,
            variable=self.direct_report_count,
            values=[str(value) for value in range(1, MAX_DIRECT_REPORT_COUNT + 1)],
            width=76,
            height=32,
            fg_color=COLORS["surface"],
            button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
            text_color=COLORS["text"],
            dropdown_fg_color=COLORS["surface"],
        )
        self.direct_report_count_menu.pack(side="left")
        self._attach_tooltip(
            self.direct_report_count_menu,
            "엑셀 희망진출국가와 추천 국가를 합쳐 만들 수출시장 분석보고서 총 개수입니다.",
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
        if self.direct_report_count_menu is not None:
            self.direct_report_count_menu.configure(
                state="normal" if self.report_mode.get() == "recommend" and self.recommend_then_direct.get() else "disabled"
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

    def _selected_direct_report_count(self) -> int:
        try:
            value = int(self.direct_report_count.get())
        except ValueError:
            return DEFAULT_DIRECT_REPORT_COUNT

        return max(1, min(MAX_DIRECT_REPORT_COUNT, value))

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
        self.progress.set("0 / 0 완료")
        self.progress_percent.set(0)
        self.total_count.set("0")
        self.running_count.set("0")
        self.completed_count.set("0")
        self.failed_count.set("0")
        self._row_status_by_key.clear()
        self._progress_row_widgets.clear()
        self._row_start_times.clear()
        self._running_keys.clear()
        self._child_keys_by_parent.clear()
        self._child_status_by_key.clear()
        self._child_state_by_key.clear()
        self._child_expanded_parents.clear()
        self._deferred_row_payloads.clear()
        self._deferred_task_payloads.clear()
        self._cancel_table_build()
        self._show_progress_empty_state()
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
        if self.direct_report_count_menu is not None:
            enabled = (
                not running
                and self.report_mode.get() == "recommend"
                and self.recommend_then_direct.get()
            )
            self.direct_report_count_menu.configure(state="normal" if enabled else "disabled")
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
        try:
            # 줄 수가 무한정 늘면 Text 위젯 갱신이 점점 느려진다(특히 윈도우).
            line_count = int(self.log_text.index("end-1c").split(".")[0])
            if line_count > _LOG_MAX_LINES:
                self.log_text.delete("1.0", f"{line_count - _LOG_MAX_LINES + 1}.0")
        except Exception:
            pass
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
                direct_report_count=self._selected_direct_report_count(),
                auto_retry_enabled=self._auto_retry_enabled,
                filename_pattern=self.active_filename_pattern,
                report_mode=self.report_mode.get(),
                recommend_then_direct=self.report_mode.get() == "recommend" and self.recommend_then_direct.get(),
                status_callback=lambda message: self.events.put(("status", message)),
                progress_callback=lambda data: self.events.put(("progress", data)),
                stop_requested=lambda: self.stop_requested,
                force_stop_requested=lambda: self.force_stop_requested,
                row_status_callback=lambda data: self.events.put(("row_status", data)),
                # 워커 스레드가 row dict 를 계속 수정하므로, GUI 가 읽을 스냅샷을 복사해 전달한다.
                rows_ready_callback=lambda rows: self.events.put(("rows_init", [dict(row) for row in rows])),
                task_status_callback=lambda data: self.events.put(("task_status", data)),
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _poll_events(self) -> None:
        processed = 0
        try:
            # 상한 없이 큐를 비울 때까지 돌면, 이벤트가 계속 쌓이는 동안 메인 스레드가
            # 루프에서 빠져나오지 못해 UI 가 멈출 수 있다(병렬 세션일수록 위험).
            while processed < _MAX_EVENTS_PER_POLL:
                event, payload = self.events.get_nowait()
                processed += 1
                if event == "status":
                    self.status.set(str(payload))
                    self._append_log(str(payload), self._log_tag(str(payload)))
                elif event == "progress":
                    self._handle_progress(payload)
                elif event == "done":
                    self._handle_done(payload)
                elif event == "rows_init":
                    self._init_progress_table(payload)
                elif event == "row_status":
                    self._handle_row_status(payload)
                elif event == "task_status":
                    self._handle_task_status(payload)
                elif event == "error":
                    self.status.set("오류 발생")
                    self._set_status_badge("danger")
                    self._append_log(f"오류 발생: {payload}", "danger")
                    self._set_running_state(False)
                    messagebox.showerror("오류", str(payload))
        except queue.Empty:
            pass
        # 상한까지 처리했다면 백로그가 남았을 수 있으니 짧게 양보 후 바로 이어서 처리
        self.after(20 if processed >= _MAX_EVENTS_PER_POLL else 200, self._poll_events)

    def _handle_progress(self, payload: object) -> None:
        data = payload if isinstance(payload, dict) else {}
        current = int(data.get("current", 0) or 0)
        total = int(data.get("total", 0) or 0)
        success = int(data.get("success", 0) or 0)
        failed = int(data.get("failed", 0) or 0)
        percent = (current / total * 100) if total else 0

        self.total_count.set(str(total))
        self.completed_count.set(str(success))
        self.failed_count.set(str(failed))
        self.progress.set(f"{current} / {total} 완료")
        self.progress_percent.set(percent)
        self._set_progress(percent)
        self._update_summary_counts()
        if data.get("status"):
            status = str(data["status"])
            self.status.set(status)
            self._append_log(status, self._log_tag(status))

    def _mode_label(self, report_mode: str, recommend_then_direct: bool) -> str:
        if report_mode == "recommend" and recommend_then_direct:
            return "추천→직접"
        if report_mode == "recommend":
            return "추천"
        return "직접"

    def _split_countries(self, value: object, limit: int | None = None) -> list[str]:
        # automation 의 분리 규칙(구분자/중복 제거/정규화)을 그대로 사용해
        # 사전 생성하는 자식 작업 목록이 실제 처리 목록과 항상 일치하게 한다.
        countries = split_country_values(value)
        if limit is not None:
            return countries[:limit]
        return countries

    def _direct_report_count_for_row(self, row: dict) -> int:
        try:
            return max(1, int(row.get("direct_report_count", 1) or 1))
        except (TypeError, ValueError):
            return 1

    def _init_queued_child_tasks(self, parent_key: str, row: dict) -> None:
        report_mode = str(row.get("report_mode", ""))
        recommend_then_direct = bool(row.get("recommend_then_direct", False))
        if report_mode == "direct":
            countries = self._split_countries(row.get("target_country", ""), self._direct_report_count_for_row(row))
        elif report_mode == "recommend" and recommend_then_direct:
            self._register_child_task(parent_key, "recommend", "", "처리 안됨")
            countries = self._split_countries(
                row.get("final_target_countries", "") or row.get("target_country", ""),
                self._direct_report_count_for_row(row),
            )
        else:
            countries = []

        for country in countries:
            self._register_child_task(parent_key, "direct", country, "처리 안됨")
        # 위젯은 만들지 않고(숨김 상태라 불필요) 토글 라벨만 갱신한다.
        # 자식 행 위젯은 실제로 보여질 때 _ensure_child_widgets 가 만든다.
        self._refresh_child_visibility(parent_key)

    def _register_child_task(self, parent_key: str, task_type: str, country: str, status: str) -> None:
        """자식 작업의 키/상태만 등록한다(위젯 생성 없음)."""
        child_key = f"{parent_key}:{task_type}:{country}"
        if child_key not in self._child_keys_by_parent.setdefault(parent_key, []):
            self._child_keys_by_parent[parent_key].append(child_key)
        self._child_status_by_key[child_key] = status
        self._child_state_by_key.setdefault(
            child_key, {"status": status, "elapsed_str": "—", "saved_files": []}
        )

    def _elapsed_str(self, key: str, ts: float) -> str:
        if key not in self._row_start_times:
            return "—"
        elapsed = int(ts - self._row_start_times[key])
        m, s = divmod(elapsed, 60)
        return f"{m}m {s:02d}s" if m else f"{s}s"

    def _tick_elapsed(self) -> None:
        now = datetime.now().timestamp()
        for key in list(self._running_keys):
            widgets = self._progress_row_widgets.get(key)
            start = self._row_start_times.get(key)
            if widgets is None or start is None:
                continue
            elapsed = int(now - start)
            m, s = divmod(elapsed, 60)
            elapsed_str = f"{m}m {s:02d}s" if m else f"{s}s"
            try:
                widgets["elapsed_lbl"].configure(text=elapsed_str)
            except Exception:
                pass
        self.after(1000, self._tick_elapsed)

    def _init_progress_table(self, payload: object) -> None:
        body = self._progress_body
        if body is None:
            return
        rows = payload if isinstance(payload, list) else []
        self._cancel_table_build()
        for w in body.winfo_children():
            w.destroy()
        self._progress_row_widgets.clear()
        self._row_start_times.clear()
        self._running_keys.clear()
        self._row_status_by_key.clear()
        self._child_keys_by_parent.clear()
        self._child_status_by_key.clear()
        self._child_state_by_key.clear()
        self._child_expanded_parents.clear()
        self._deferred_row_payloads.clear()
        self._deferred_task_payloads.clear()
        # 파괴된 옛 행 프레임들을 _table_frames 에서 정리(재실행 반복 시 누적 방지)
        self._apply_table_column_layout()
        if not rows:
            self._show_progress_empty_state()
            return
        # 1단계: 위젯 없이 행 키/상태만 먼저 등록한다.
        # (요약 카운트가 즉시 정확해지고, 빌드 중 도착한 이벤트를 버퍼링할지 판정할 수 있다)
        specs: list[tuple[str, dict]] = []
        for row in rows:
            row_index = int(row.get("row_index", 0))
            key = str(row.get("ui_key", "") or f"row:{row_index}")
            self._row_status_by_key[key] = "처리 안됨"
            specs.append((key, row))
        self._update_summary_counts()
        # 2단계: 행 위젯은 after 청크로 나눠 생성한다. 한 번에 모두 만들면 행이 많은
        # 엑셀에서 메인 스레드가 수 초간 멈춰 윈도우에서 '응답 없음'이 뜬다.
        self._table_build_specs = specs
        self._table_build_index = 0
        self._build_table_chunk()

    def _cancel_table_build(self) -> None:
        if self._table_build_job is not None:
            try:
                self.after_cancel(self._table_build_job)
            except Exception:
                pass
            self._table_build_job = None
        self._table_build_specs = []
        self._table_build_index = 0

    def _build_table_chunk(self) -> None:
        self._table_build_job = None
        body = self._progress_body
        if body is None:
            return
        specs = self._table_build_specs
        index = self._table_build_index
        end = min(index + _TABLE_BUILD_CHUNK_ROWS, len(specs))
        for key, row in specs[index:end]:
            row_index = int(row.get("row_index", 0))
            mode_label = self._mode_label(
                str(row.get("report_mode", "")),
                bool(row.get("recommend_then_direct", False)),
            )
            widgets = self._make_data_row(
                body,
                row_index=row_index,
                product=str(row.get("product_name", ""))[:35],
                hs=str(row.get("hs_code", "")),
                # 희망 국가 없음(플레이스홀더 '-' 입력 등)은 빈 값으로 정규화되므로 '—' 로 표시
                country=str(row.get("target_country", ""))[:20] or "—",
                mode_lbl=mode_label,
            )
            self._progress_row_widgets[key] = widgets
            self._init_queued_child_tasks(key, row)
            # 이 행 위젯이 만들어지기 전에 도착해 보관해 둔 이벤트를 즉시 반영한다.
            deferred_row = self._deferred_row_payloads.pop(key, None)
            if deferred_row is not None:
                self._handle_row_status(deferred_row)
            for task_payload in self._deferred_task_payloads.pop(key, {}).values():
                self._upsert_child_task(
                    key,
                    str(task_payload.get("task_type", "")),
                    str(task_payload.get("country", "")),
                    str(task_payload.get("status", "")),
                    float(task_payload.get("ts", 0)),
                    [str(p) for p in task_payload.get("saved_files", [])],
                )
        self._table_build_index = end
        if end < len(specs):
            self._table_build_job = self.after(_TABLE_BUILD_CHUNK_DELAY_MS, self._build_table_chunk)
            return
        self._table_build_specs = []
        self._table_build_index = 0
        try:
            canvas = getattr(body, "_parent_canvas", None)
            if canvas is not None:
                canvas.yview_moveto(0)
        except Exception:
            pass
        self._update_summary_counts()

    def _handle_row_status(self, payload: object) -> None:
        data = payload if isinstance(payload, dict) else {}
        row_index = int(data.get("row_index", 0))
        status = str(data.get("status", ""))
        session_id = int(data.get("session_id", 1))
        ts = float(data.get("ts", 0))
        key = str(data.get("ui_key", "") or f"row:{row_index}")
        widgets = self._progress_row_widgets.get(key)
        if widgets is None:
            # 행 위젯이 청크 빌드 대기 중이면 마지막 상태를 보관했다가 생성 직후 반영한다.
            if key in self._row_status_by_key:
                self._deferred_row_payloads[key] = dict(data)
                self._row_status_by_key[key] = status
                self._update_summary_counts()
            return
        if status == "처리 중":
            self._running_keys.add(key)
            self._row_start_times[key] = ts
            elapsed_str = "0s"
        else:
            self._running_keys.discard(key)
            elapsed_str = self._elapsed_str(key, ts)
        self._row_status_by_key[key] = status
        saved_files = data.get("saved_files", [])
        if not isinstance(saved_files, list):
            saved_files = []
        try:
            self._configure_status_pill(widgets["status_lbl"], status)
            # session_id=0 은 세션에 배정되지 못한 채 종료된 행(병렬 세션 중단 등).
            widgets["session_lbl"].configure(text=f"S{session_id}" if session_id > 0 else "—")
            widgets["elapsed_lbl"].configure(text=elapsed_str)
            self._update_file_button(widgets, [str(path) for path in saved_files])
        except Exception:
            pass
        self._update_summary_counts()

    def _handle_task_status(self, payload: object) -> None:
        data = payload if isinstance(payload, dict) else {}
        row_index = int(data.get("row_index", 0))
        task_type = str(data.get("task_type", ""))
        country = str(data.get("country", ""))
        status = str(data.get("status", ""))
        ts = float(data.get("ts", 0))
        saved_files = data.get("saved_files", [])
        if not isinstance(saved_files, list):
            saved_files = []
        parent_key = str(data.get("ui_key", "") or f"row:{row_index}")
        self._upsert_child_task(parent_key, task_type, country, status, ts, [str(path) for path in saved_files])

    def _upsert_child_task(
        self,
        parent_key: str,
        task_type: str,
        country: str,
        status: str,
        ts: float,
        saved_files: list[str] | None = None,
    ) -> None:
        parent_widgets = self._progress_row_widgets.get(parent_key)
        child_key = f"{parent_key}:{task_type}:{country}"
        if parent_widgets is None:
            # 행 위젯이 아직 청크 빌드 대기 중이면 마지막 페이로드만 보관해 두고,
            # 행 생성 직후 _build_table_chunk 가 다시 흘려보낸다.
            if parent_key in self._row_status_by_key:
                self._deferred_task_payloads.setdefault(parent_key, {})[child_key] = {
                    "task_type": task_type,
                    "country": country,
                    "status": status,
                    "ts": ts,
                    "saved_files": list(saved_files or []),
                }
            return
        if status == "처리 중":
            self._row_start_times[child_key] = ts
            self._running_keys.add(child_key)
            elapsed_str = "0s"
        else:
            self._running_keys.discard(child_key)
            if status == "처리 안됨":
                # 자동 재시도 시 자식 작업이 '처리 안됨'으로 재초기화된다.
                # 이전 시도의 시작 시각이 남아 있으면 경과시간이 잘못 표시되므로 비운다.
                self._row_start_times.pop(child_key, None)
                elapsed_str = "—"
            else:
                elapsed_str = self._elapsed_str(child_key, ts)
        if child_key not in self._child_keys_by_parent.setdefault(parent_key, []):
            self._child_keys_by_parent[parent_key].append(child_key)
        self._child_status_by_key[child_key] = status
        # 숨김 자식은 위젯을 만들지 않으므로, 나중에 보일 때 적용할 표시 상태를 보관한다.
        self._child_state_by_key[child_key] = {
            "status": status,
            "elapsed_str": elapsed_str,
            "saved_files": list(saved_files or []),
        }
        child_widgets = self._progress_row_widgets.get(child_key)
        if child_widgets is not None:
            try:
                self._configure_status_pill(child_widgets["status_lbl"], status)
                child_widgets["elapsed_lbl"].configure(text=elapsed_str)
                self._update_file_button(child_widgets, saved_files or [])
            except Exception:
                pass
        self._refresh_child_visibility(parent_key)

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
