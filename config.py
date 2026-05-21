from __future__ import annotations

import sys
from pathlib import Path


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent


KOTRA_REPORT_URL = "https://kotra.or.kr/mutugpt/export-assistant/report"

DEFAULT_INPUT_EXCEL = BASE_DIR / "input_template.xlsx"
DEFAULT_DOWNLOAD_DIR = BASE_DIR / "downloads"
DEFAULT_LOG_DIR = BASE_DIR / "logs"
DEFAULT_STATE_PATH = BASE_DIR / "state.json"

TIMEOUT_MINUTES = 15
TIMEOUT_MS = TIMEOUT_MINUTES * 60 * 1000
PAGE_LOAD_TIMEOUT_MS = 120 * 1000
ELEMENT_TIMEOUT_MS = 60 * 1000
GENERATION_RETRY_COUNT = 2

# 개발 초기 기본값은 브라우저를 표시하는 모드입니다.
DEFAULT_HEADLESS = False

DEFAULT_CONFIG = {
    "headless": DEFAULT_HEADLESS,
    "timeout_minutes": TIMEOUT_MINUTES,
    "page_load_timeout_ms": PAGE_LOAD_TIMEOUT_MS,
    "element_timeout_ms": ELEMENT_TIMEOUT_MS,
    "download_dir": str(DEFAULT_DOWNLOAD_DIR),
    "log_dir": str(DEFAULT_LOG_DIR),
    "state_path": str(DEFAULT_STATE_PATH),
    "retry_count": GENERATION_RETRY_COUNT,
    "use_storage_state": False,
    "save_storage_state": False,
}


APP_CREDITS = {
    "app_name": "KOTRA 수출시장 분석보고서 자동생성기",
    "version": "1.0.0",
    "purpose": (
        "수출시장 분석보고서 생성 과정의 반복 입력, "
        "보고서 생성 대기, 파일 다운로드 업무를 자동화하기 위해 제작된 "
        "내부 업무 보조 도구입니다."
    ),
    "developed_by": "기현명",
    "role": "KOTRA AI데이터팀 청년인턴",
    "tech_stack": "Python, Playwright, CustomTkinter",
    "development_support": "OpenAI Codex",
    "contact": "kiwizcloud@kotra.or.kr",
    "disclaimer": (
        "본 프로그램은 내부 업무 효율화를 위한 보조 도구이며, "
        "KOTRA의 공식 대외 서비스가 아닙니다."
    ),
    "copyright": "© 2026 KOTRA. Internal use only.",
}
