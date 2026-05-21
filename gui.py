from __future__ import annotations

import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from automation import run_automation
from config import APP_CREDITS, BASE_DIR, DEFAULT_DOWNLOAD_DIR, DEFAULT_LOG_DIR, DEFAULT_STATE_PATH
from template import create_input_template


class KotraReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("KOTRA 보고서 자동 생성기")
        self._configure_style()
        self._build_menu()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = max(820, min(1040, screen_width - 120))
        height = max(560, min(760, screen_height - 140))
        self.geometry(f"{width}x{height}")
        self.minsize(820, 560)
        self.resizable(True, True)

        self.input_path = tk.StringVar(value=str(BASE_DIR / "input.xlsx"))
        self.download_dir = tk.StringVar(value=str(DEFAULT_DOWNLOAD_DIR))
        self.run_mode = tk.StringVar(value="all")
        self.background = tk.BooleanVar(value=False)
        self.use_session = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="대기 중")
        self.progress = tk.StringVar(value="0 / 0")
        self.progress_percent = tk.DoubleVar(value=0)
        self.success = tk.StringVar(value="0건")
        self.failed = tk.StringVar(value="0건")
        self.status_label: ttk.Label | None = None
        self.log_text: ScrolledText | None = None
        self.info_label: tk.Label | None = None
        self.start_button: ttk.Button | None = None
        self.stop_button: ttk.Button | None = None
        self.force_stop_button: ttk.Button | None = None

        self.stop_requested = False
        self.force_stop_requested = False
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_ui()
        self.after(200, self._poll_events)

    def _configure_style(self) -> None:
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=11)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=11)

        self.configure(background="#f7f8fa")
        style = ttk.Style(self)
        style.configure("TFrame", background="#f7f8fa")
        style.configure("TLabel", background="#f7f8fa", foreground="#202124")
        style.configure("Muted.TLabel", background="#f7f8fa", foreground="#6b7280")
        style.configure("Title.TLabel", background="#f7f8fa", foreground="#202124", font=("", 22, "bold"))
        style.configure("Section.TLabelframe", background="#f7f8fa", padding=12)
        style.configure("Section.TLabelframe.Label", background="#f7f8fa", foreground="#374151", font=("", 12, "bold"))
        style.configure("Primary.TButton", padding=(18, 7))
        style.configure("Action.TButton", padding=(14, 7))

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self)
        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="프로그램 정보", command=self._show_about)
        menu_bar.add_cascade(label="도움말", menu=help_menu)
        self.config(menu=menu_bar)

    def _build_ui(self) -> None:
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0, background="#f7f8fa")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        root = ttk.Frame(canvas, padding=24)
        window_id = canvas.create_window((0, 0), window=root, anchor="nw")

        root.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: self._resize_canvas_content(canvas, window_id, event.width))
        canvas.bind_all("<MouseWheel>", lambda event: self._on_mousewheel(canvas, event))

        header_frame = ttk.Frame(root)
        header_frame.grid(row=0, column=0, sticky="we", pady=(0, 16))
        ttk.Label(header_frame, text=APP_CREDITS["app_name"], style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.info_label = tk.Label(
            header_frame,
            text="ⓘ",
            width=2,
            cursor="hand2",
            bg="#f7f8fa",
            fg="#6b7280",
            font=("", 18),
        )
        self.info_label.grid(row=0, column=1, sticky="ne", padx=(12, 0), pady=(2, 0))
        self.info_label.bind("<Button-1>", lambda _event: self._show_about())
        self.info_label.bind("<Enter>", lambda _event: self._set_info_hover(True))
        self.info_label.bind("<Leave>", lambda _event: self._set_info_hover(False))
        ttk.Label(header_frame, text="엑셀 데이터를 바탕으로 KOTRA 수출 시장 분석 보고서를 자동 생성하고 저장합니다.", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        header_frame.columnconfigure(0, weight=1)

        ttk.Separator(root).grid(row=1, column=0, sticky="we", pady=(0, 14))

        file_frame = ttk.LabelFrame(root, text="파일", style="Section.TLabelframe")
        file_frame.grid(row=2, column=0, sticky="we", pady=(0, 12))
        ttk.Label(file_frame, text="엑셀 파일").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(file_frame, textvariable=self.input_path, width=20).grid(row=0, column=1, sticky="we", pady=8)
        ttk.Button(file_frame, text="찾기", style="Action.TButton", command=self._choose_input).grid(row=0, column=2, padx=10, pady=8)

        ttk.Label(file_frame, text="저장 폴더").grid(row=1, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(file_frame, textvariable=self.download_dir, width=20).grid(row=1, column=1, sticky="we", pady=8)
        ttk.Button(file_frame, text="찾기", style="Action.TButton", command=self._choose_download_dir).grid(row=1, column=2, padx=10, pady=8)
        file_frame.columnconfigure(1, weight=1)

        option_frame = ttk.LabelFrame(root, text="실행 옵션", style="Section.TLabelframe")
        option_frame.grid(row=3, column=0, sticky="we", pady=(0, 12))
        ttk.Radiobutton(option_frame, text="전체 실행", variable=self.run_mode, value="all").grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        ttk.Radiobutton(option_frame, text="실패 행만 재시도", variable=self.run_mode, value="retry_failed").grid(row=1, column=0, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(option_frame, text="백그라운드 실행", variable=self.background).grid(row=0, column=1, sticky="w", padx=24, pady=(8, 2))
        ttk.Checkbutton(option_frame, text="브라우저 세션 저장 사용(필요 시)", variable=self.use_session).grid(row=1, column=1, sticky="w", padx=24, pady=2)
        ttk.Label(option_frame, text="실패 행 재시도는 logs/failed_rows.xlsx 기준입니다.", style="Muted.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 8))

        button_row = ttk.Frame(root)
        button_row.grid(row=4, column=0, sticky="w", pady=(0, 10))
        self.start_button = ttk.Button(button_row, text="시작", style="Primary.TButton", command=self._start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(button_row, text="중지", style="Action.TButton", command=self._stop, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))
        self.force_stop_button = ttk.Button(button_row, text="강제종료", style="Action.TButton", command=self._force_stop, state="disabled")
        self.force_stop_button.pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="예시 템플릿 생성", style="Action.TButton", command=self._create_template).pack(side="left", padx=(8, 0))

        ttk.Label(root, text="중지: 현재 행까지 실행 후 멈춤\n강제종료: 즉시 중단", style="Muted.TLabel").grid(row=5, column=0, sticky="w", pady=(0, 12))

        progress_frame = ttk.LabelFrame(root, text="진행 상황", style="Section.TLabelframe")
        progress_frame.grid(row=6, column=0, sticky="we", pady=(0, 12))
        ttk.Label(progress_frame, text="진행률").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        ttk.Label(progress_frame, textvariable=self.progress).grid(row=0, column=1, sticky="w", pady=6)
        ttk.Progressbar(progress_frame, variable=self.progress_percent, maximum=100).grid(row=0, column=2, sticky="we", padx=10, pady=6)

        ttk.Label(progress_frame, text="현재 상태").grid(row=1, column=0, sticky="nw", padx=10, pady=6)
        self.status_label = ttk.Label(progress_frame, textvariable=self.status, wraplength=560)
        self.status_label.grid(row=1, column=1, columnspan=2, sticky="we", pady=6, padx=(0, 10))

        ttk.Label(progress_frame, text="처리 결과").grid(row=2, column=0, sticky="w", padx=10, pady=6)
        stats_frame = ttk.Frame(progress_frame)
        stats_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=6)
        ttk.Label(stats_frame, text="성공").pack(side="left")
        ttk.Label(stats_frame, textvariable=self.success).pack(side="left", padx=(6, 18))
        ttk.Label(stats_frame, text="실패").pack(side="left")
        ttk.Label(stats_frame, textvariable=self.failed).pack(side="left", padx=(6, 0))
        progress_frame.columnconfigure(2, weight=1)

        log_frame = ttk.LabelFrame(root, text="실행 로그", style="Section.TLabelframe")
        log_frame.grid(row=7, column=0, sticky="nsew")
        self.log_text = ScrolledText(log_frame, height=8, wrap="word", state="disabled", relief="flat", borderwidth=0, background="#fbfbfc")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

        root.columnconfigure(0, weight=1)
        root.rowconfigure(7, weight=1)

    def _resize_canvas_content(self, canvas: tk.Canvas, window_id: int, width: int) -> None:
        content_width = max(width, self.minsize()[0])
        canvas.itemconfigure(window_id, width=content_width)
        if self.status_label is not None:
            self.status_label.configure(wraplength=max(260, content_width - 210))

    def _on_mousewheel(self, canvas: tk.Canvas, event: tk.Event) -> None:
        if event.delta == 0:
            return
        step = -1 if event.delta > 0 else 1
        canvas.yview_scroll(step, "units")

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

    def _set_info_hover(self, hovered: bool) -> None:
        if self.info_label is None:
            return
        self.info_label.configure(fg="#2563eb" if hovered else "#6b7280")

    def _show_about(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("프로그램 정보")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(background="#f7f8fa")

        body = ttk.Frame(dialog, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=APP_CREDITS["app_name"], style="Title.TLabel").pack(anchor="w")
        ttk.Label(body, text=f"Version {APP_CREDITS['version']}", style="Muted.TLabel").pack(anchor="w", pady=(4, 10))
        ttk.Label(body, text=APP_CREDITS["purpose"], style="Muted.TLabel", wraplength=460).pack(anchor="w", pady=(0, 18))

        info = ttk.Frame(body)
        info.pack(fill="x", pady=(0, 18))
        rows = [
            ("제작", APP_CREDITS["developed_by"]),
            ("소속/역할", APP_CREDITS["role"]),
            ("기술 스택", APP_CREDITS["tech_stack"]),
            ("개발 지원", APP_CREDITS["development_support"]),
            ("문의", APP_CREDITS["contact"]),
        ]
        for index, (label, value) in enumerate(rows):
            ttk.Label(info, text=label, width=12, style="Muted.TLabel").grid(row=index, column=0, sticky="w", pady=3)
            ttk.Label(info, text=value, wraplength=360).grid(row=index, column=1, sticky="w", pady=3)

        ttk.Separator(body).pack(fill="x", pady=(0, 12))
        ttk.Label(body, text=APP_CREDITS["disclaimer"], style="Muted.TLabel", wraplength=460).pack(anchor="w", pady=(0, 8))
        ttk.Label(body, text=APP_CREDITS["copyright"], style="Muted.TLabel").pack(anchor="w")
        ttk.Button(body, text="확인", style="Action.TButton", command=dialog.destroy).pack(anchor="e", pady=(20, 0))

        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dialog.grab_set()

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("실행 중", "이미 작업이 실행 중입니다.")
            return

        input_path = Path(self.input_path.get())
        if self.run_mode.get() == "all" and not input_path.exists():
            messagebox.showerror("오류", f"엑셀 파일을 찾을 수 없습니다.\n{input_path}")
            return

        self.stop_requested = False
        self.force_stop_requested = False
        self.progress.set("0 / 0")
        self.progress_percent.set(0)
        self.success.set("0건")
        self.failed.set("0건")
        self.status.set("시작 준비 중")
        self._clear_log()
        self._append_log("작업을 시작합니다.")
        self._set_running_state(True)
        self.worker = threading.Thread(target=self._run_worker, daemon=True)
        self.worker.start()

    def _stop(self) -> None:
        self.stop_requested = True
        self.status.set("중지 요청됨: 현재 행 처리 후 멈춥니다.")
        self._append_log("중지 요청됨: 현재 행 처리 후 멈춥니다.")

    def _force_stop(self) -> None:
        self.force_stop_requested = True
        self.stop_requested = True
        self.status.set("강제종료 요청됨: 현재 작업을 즉시 중단합니다.")
        self._append_log("강제종료 요청됨: 현재 작업을 즉시 중단합니다.")

    def _set_running_state(self, running: bool) -> None:
        if self.start_button is not None:
            self.start_button.configure(state="disabled" if running else "normal")
        if self.stop_button is not None:
            self.stop_button.configure(state="normal" if running else "disabled")
        if self.force_stop_button is not None:
            self.force_stop_button.configure(state="normal" if running else "disabled")

    def _clear_log(self) -> None:
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _append_log(self, message: str) -> None:
        if self.log_text is None:
            return
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
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
                retry_failed_only=self.run_mode.get() == "retry_failed",
                wait_for_manual_login=False,
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
                    self._append_log(str(payload))
                elif event == "progress":
                    data = payload if isinstance(payload, dict) else {}
                    current = int(data.get("current", 0) or 0)
                    total = int(data.get("total", 0) or 0)
                    self.progress.set(f"{current} / {total}")
                    self.progress_percent.set((current / total * 100) if total else 0)
                    self.success.set(f"{data.get('success', 0)}건")
                    self.failed.set(f"{data.get('failed', 0)}건")
                    if data.get("status"):
                        self.status.set(str(data["status"]))
                        self._append_log(str(data["status"]))
                elif event == "done":
                    result = payload if isinstance(payload, dict) else {}
                    self.status.set("완료")
                    self._append_log("작업이 완료되었습니다.")
                    self._set_running_state(False)
                    messagebox.showinfo(
                        "완료",
                        (
                            f"전체 {result.get('total', 0)}건\n"
                            f"성공 {result.get('success', 0)}건\n"
                            f"실패 {result.get('failed', 0)}건"
                        ),
                    )
                elif event == "error":
                    self.status.set("오류 발생")
                    self._append_log(f"오류 발생: {payload}")
                    self._set_running_state(False)
                    messagebox.showerror("오류", str(payload))
        except queue.Empty:
            pass
        self.after(200, self._poll_events)


def run_gui() -> None:
    app = KotraReportApp()
    app.mainloop()
