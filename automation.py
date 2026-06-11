from __future__ import annotations

import queue
import re
import importlib.util
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from compat import ensure_stdlib_selectors

ensure_stdlib_selectors()

import pandas as pd
from playwright.sync_api import Error as PlaywrightError, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

from config import (
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_DIRECT_REPORT_COUNT,
    DEFAULT_HEADLESS,
    DEFAULT_LOG_DIR,
    DEFAULT_ROW_RETRY_COUNT,
    DEFAULT_STATE_PATH,
    ELEMENT_TIMEOUT_MS,
    GENERATION_RETRY_COUNT,
    KOTRA_REPORT_URL,
    MAX_DIRECT_REPORT_COUNT,
    MAX_PARALLEL_SESSIONS,
    PAGE_LOAD_TIMEOUT_MS,
    TIMEOUT_MS,
)
from field_mapping import (
    DIRECT_FIELD_MAPPING,
    EXPORT_EXPERIENCE_CATEGORY_MAP,
    EXPORT_SCALE_CATEGORY_MAP,
    FIELD_MAPPING,
    RECOMMEND_FIELD_MAPPING,
    SOURCE_COLUMN_ALIASES,
)
from logger import log_failed_row, log_success_row


def _load_app_selectors() -> dict[str, str]:
    selectors_path = Path(__file__).resolve().parent / "site_selectors.py"
    spec = importlib.util.spec_from_file_location("kotra_app_selectors", selectors_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"selector 설정 파일을 불러올 수 없습니다: {selectors_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.SELECTORS


SELECTORS = _load_app_selectors()

StatusCallback = Callable[[str], None]
ProgressCallback = Callable[[dict[str, Any]], None]
RowStatusCallback = Callable[[dict[str, Any]], None]
TaskStatusCallback = Callable[[dict[str, Any]], None]
StopFlag = Callable[[], bool]
RetryFlag = Callable[[], bool]

REPORT_MODE_DIRECT = "direct"
REPORT_MODE_RECOMMEND = "recommend"
REPORT_MODE_LABELS = {
    REPORT_MODE_DIRECT: "수출시장 분석 보고서 생성",
    REPORT_MODE_RECOMMEND: "유망 시장 추천 보고서 생성",
}
PROCESSING_STATUS_FILENAME = "processing_status.xlsx"
REPORT_TASKS_FILENAME = "report_tasks.xlsx"
SOURCE_FILE_COLUMN = "원본파일"
STATUS_COLUMN = "처리상태"
STATUS_AT_COLUMN = "처리일시"
SAVED_FILE_COLUMN = "저장파일"
ERROR_COLUMN = "오류메시지"
STATUS_COLUMNS = [
    STATUS_COLUMN,
    STATUS_AT_COLUMN,
    SAVED_FILE_COLUMN,
    ERROR_COLUMN,
]
PROCESSING_STATUS_COLUMNS = [
    SOURCE_FILE_COLUMN,
    "report_mode",
    "recommend_then_direct",
    "direct_report_count",
    "row_index",
    "hs_code",
    "product_name",
    "export_scale",
    "export_experience",
    "target_country",
    "excluded_countries",
    "recommended_countries",
    "final_target_countries",
    "recommendation_report_file",
    "direct_report_files",
    *STATUS_COLUMNS,
]
STATUS_PENDING = "처리 안됨"
STATUS_RUNNING = "처리 중"
STATUS_RETRY_PENDING = "자동 재시도 대기"
STATUS_SUCCESS = "처리완료"
STATUS_FAILED = "처리실패"
TASK_TYPE_RECOMMEND = "recommend"
TASK_TYPE_DIRECT = "direct"
REPORT_TASK_COLUMNS = [
    SOURCE_FILE_COLUMN,
    "report_mode",
    "recommend_then_direct",
    "direct_report_count",
    "row_index",
    "hs_code",
    "product_name",
    "target_country",
    "excluded_countries",
    "task_type",
    "country",
    "status",
    "saved_file",
    "recommended_countries",
    "final_target_countries",
    "error_message",
    "updated_at",
]
REPORT_TASKS_LOCK = threading.Lock()
PARALLEL_WAIT_SUMMARY_INTERVAL_SECONDS = 30
PARALLEL_SESSION_START_DELAY_SECONDS = 3
DEFAULT_INITIAL_FORM_WAIT_SECONDS = 10
PARALLEL_INITIAL_FORM_WAIT_SECONDS = 15
GENERATION_STATUS_POLL_INTERVAL_MS = 3_000
# 보고서 생성 스트리밍 API 경로. 이 경로의 POST가 끊기면 생성이 죽은 것으로 본다.
GENERATION_API_URL_MARKER = "/api/export-assistant/"
# 생성 요청 중단 감지 후, 사이트가 인식 가능한 화면(초기화면/오류/중단)으로 전환되길 기다리는 유예시간.
GENERATION_ABORT_GRACE_SECONDS = 30
# 페이지 전체 텍스트가 이 시간 동안 전혀 변하지 않으면 생성이 멈춘 것으로 판정한다.
GENERATION_STALL_SECONDS = 240
SCREENSHOT_TIMEOUT_MS = 15_000
FILENAME_PATTERN_TOKEN_LABELS = [
    ("연번", "row_index"),
    ("회사명", "company_name"),
    ("사업자번호", "business_number"),
    ("HS CODE", "hs_code"),
    ("수출품명", "product_name"),
    ("희망진출국가", "target_country"),
    ("분석제외국가", "excluded_countries"),
    ("보고서생성방식", "report_mode"),
    ("생성날짜", "date"),
    ("생성시간", "time"),
    ("생성일시", "datetime"),
    ("년", "year"),
    ("월", "month"),
    ("일", "day"),
    ("시", "hour"),
    ("분", "minute"),
    ("초", "second"),
    ("사이트 기본 파일명", "site_filename"),
]
FILENAME_PATTERN_TOKENS = {token for _label, token in FILENAME_PATTERN_TOKEN_LABELS}
FILENAME_PATTERN_TOKEN_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")
FILENAME_FORBIDDEN_CHARS_RE = re.compile(r'[\\/:*?"<>|]')
FILENAME_CONTROL_CHARS_RE = re.compile(r"[\x00-\x1f]")
WINDOWS_RESERVED_FILENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}
MAX_SAFE_FILENAME_LENGTH = 240


class GenerationError(RuntimeError):
    def __init__(self, message: str, artifacts: str = "") -> None:
        super().__init__(message)
        self.artifacts = artifacts


class AutomationAborted(RuntimeError):
    pass


@dataclass
class RowProcessResult:
    saved_files: list[Path]
    recommendation_report_file: Path | None = None
    direct_report_files: list[Path] | None = None
    recommended_countries: list[str] | None = None
    final_target_countries: list[str] | None = None

    def saved_files_text(self) -> str:
        return "; ".join(str(path) for path in self.saved_files)

    def saved_file_names_text(self) -> str:
        return ", ".join(path.name for path in self.saved_files)


def launch_browser(playwright: Any, headless: bool, status_callback: StatusCallback | None = None):
    def emit(message: str) -> None:
        if status_callback:
            status_callback(message)

    headless_shell_exc: PlaywrightError | None = None
    if headless:
        try:
            emit("백그라운드 실행: Playwright Chromium headless shell을 실행합니다.")
            return playwright.chromium.launch(headless=True)
        except PlaywrightError as exc:
            headless_shell_exc = exc
            emit("Playwright headless shell 실행에 실패해 Microsoft Edge headless로 다시 시도합니다.")

    launch_options: dict[str, Any] = {"headless": headless}
    if headless:
        launch_options["args"] = ["--disable-gpu", "--no-startup-window"]

    try:
        mode = "백그라운드" if headless else "화면 표시"
        emit(f"Microsoft Edge를 {mode} 모드로 실행합니다.")
        return playwright.chromium.launch(channel="msedge", **launch_options)
    except PlaywrightError as edge_exc:
        emit("Microsoft Edge 실행에 실패해 Playwright 기본 Chromium으로 다시 시도합니다.")
        try:
            return playwright.chromium.launch(headless=headless)
        except PlaywrightError as chromium_exc:
            headless_shell_message = ""
            if headless_shell_exc is not None:
                headless_shell_message = f"- Headless shell 원본 오류: {headless_shell_exc}\n"
            raise RuntimeError(
                "브라우저를 Playwright로 실행하지 못했습니다.\n"
                f"{headless_shell_message}"
                "- Windows VM에서는 Microsoft Edge가 설치되어 있는지 확인해주세요.\n"
                "- 백그라운드 실행 오류가 계속되면 install_vm.bat를 다시 실행해 Playwright headless shell 설치를 확인해주세요.\n"
                "- macOS 등 로컬 테스트 환경에서는 Playwright Chromium이 설치되어 있는지 확인해주세요.\n"
                "- 사내 브라우저 정책이 자동화 실행 또는 새 브라우저 프로필 생성을 차단할 수 있습니다.\n"
                "- 브라우저가 실행되지만 다운로드가 실패하면 보안/DLP 정책의 다운로드 차단 여부를 확인해주세요.\n"
                f"Edge 원본 오류: {edge_exc}\n"
                f"Chromium 원본 오류: {chromium_exc}"
            ) from chromium_exc


def browser_is_closed(browser: Any) -> bool:
    try:
        return bool(browser.is_connected and not browser.is_connected())
    except Exception:
        return True


def page_is_closed(page: Any) -> bool:
    try:
        return bool(page.is_closed())
    except Exception:
        return True


def close_browser_context(context: Any | None, browser: Any | None) -> None:
    for item in (context, browser):
        if item is None:
            continue
        try:
            item.close()
        except Exception:
            pass


def create_context_page(
    browser: Any,
    *,
    use_storage_state: bool,
    state_path: Path,
) -> tuple[Any, Page, list[str]]:
    context_kwargs: dict[str, Any] = {"accept_downloads": True}
    if use_storage_state and state_path.exists():
        context_kwargs["storage_state"] = str(state_path)

    context = browser.new_context(**context_kwargs)
    page = context.new_page()
    page.set_default_navigation_timeout(PAGE_LOAD_TIMEOUT_MS)
    page.set_default_timeout(ELEMENT_TIMEOUT_MS)
    diagnostic_events = setup_page_diagnostics(page)
    ensure_generation_failure_recorder(page)
    return context, page, diagnostic_events


def write_startup_error(log_dir: str | Path, message: str) -> Path:
    path = Path(log_dir) / "startup_error.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"{datetime.now():%Y-%m-%d %H:%M:%S}\n{message}\n", encoding="utf-8")
    return path


def open_site(page: Page, timeout_ms: int = PAGE_LOAD_TIMEOUT_MS, retry_count: int = 1) -> None:
    last_error: Exception | None = None

    for attempt in range(1, retry_count + 2):
        try:
            # 일부 공공 사이트는 부가 리소스 때문에 domcontentloaded가 늦을 수 있습니다.
            # 먼저 문서 응답 시작까지만 확인하고, 실제 준비 상태는 이후 selector 대기로 판단합니다.
            page.goto(KOTRA_REPORT_URL, wait_until="commit", timeout=timeout_ms)
            return
        except PlaywrightTimeoutError as exc:
            last_error = exc
            if attempt > retry_count:
                break

    raise RuntimeError(
        f"KOTRA 페이지 접속이 {timeout_ms // 1000}초 안에 완료되지 않았습니다. "
        "브라우저에서 사이트가 열리는지 확인한 뒤 다시 실행해주세요."
    ) from last_error


def select_direct_country_analysis(page: Page) -> None:
    """
    분석 방식에서 '희망 국가 직접 분석'을 선택한다.
    """
    button = page.locator(SELECTORS["direct_analysis_button"]).first
    button.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    button.scroll_into_view_if_needed()
    button.click()


def select_recommend_market_analysis(page: Page) -> None:
    """
    분석 방식에서 '유망 시장 추천 받기'를 선택한다.
    """
    button = page.locator(SELECTORS["recommend_analysis_button"]).first
    button.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    button.scroll_into_view_if_needed()
    button.click()


def fill_form(page: Page, row_data: dict[str, Any], field_mapping: dict[str, dict[str, Any]] | None = None) -> None:
    """
    field_mapping을 기준으로 엑셀 한 행의 데이터를 사이트 입력창에 채운다.
    """
    mapping = field_mapping or DIRECT_FIELD_MAPPING
    for column_name, field_info in mapping.items():
        value = str(row_data.get(column_name, "")).strip()

        if field_info.get("required") and not value:
            raise ValueError(f"필수 입력값 누락: {column_name}")

        if not value:
            continue
        if column_name == "hs_code":
            validate_hs_code(value)

        field_type = field_info["type"]
        selector_key = field_info["selector_key"]

        if field_type == "input":
            selector = SELECTORS[selector_key]
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            locator.scroll_into_view_if_needed()
            locator.fill(value)

        elif field_type == "dropdown":
            select_dropdown(page, selector_key, value, field_info)

        elif field_type == "button_group":
            if column_name == "export_experience":
                click_export_experience(page, value)
            else:
                click_button_group_option(page, field_info, value)

        else:
            raise ValueError(f"지원하지 않는 입력 타입입니다: {field_type}")


def _dropdown_selected(dropdown, candidates: list[str]) -> bool:
    try:
        text = dropdown.inner_text(timeout=1_000).strip().replace(" ", "")
        return any(c.replace(" ", "") in text for c in candidates if c)
    except Exception:
        return False


def select_dropdown(page: Page, selector_key: str, value: str, field_info: dict[str, Any] | None = None) -> None:
    selector = SELECTORS[selector_key]
    dropdown = page.locator(selector).first
    dropdown.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    dropdown.scroll_into_view_if_needed()

    candidates = dropdown_value_candidates(value, field_info or {})
    open_dropdown(page, dropdown, candidates)

    if click_dropdown_option(page, candidates):
        return

    if select_dropdown_by_keyboard(page, dropdown, candidates[0], field_info or {}):
        if _dropdown_selected(dropdown, candidates):
            return

    open_dropdown(page, dropdown, candidates)
    if click_dropdown_option(page, candidates):
        return

    joined = ", ".join(candidates)
    raise ValueError(f"드롭다운 옵션을 찾지 못했습니다. 입력값: {value}, 시도한 값: {joined}")


def open_dropdown(page: Page, dropdown, candidates: list[str]) -> None:
    open_attempts = [
        lambda: dropdown.click(),
        lambda: click_locator_right_edge(page, dropdown),
        lambda: dropdown.locator("xpath=ancestor-or-self::*[@role='combobox' or @role='button' or self::button][1]").click(timeout=3_000),
        lambda: dropdown.press("Enter"),
        lambda: dropdown.press("Space"),
    ]

    last_error: Exception | None = None
    for attempt in open_attempts:
        try:
            attempt()
            if wait_for_any_text_visible(page, candidates, timeout_ms=2_000):
                return
        except Exception as exc:
            last_error = exc

    if last_error:
        # 옵션이 늦게 뜨는 사이트를 위해 마지막으로 한 번 더 짧게 확인합니다.
        wait_for_any_text_visible(page, candidates, timeout_ms=1_000)


def click_locator_right_edge(page: Page, locator) -> None:
    box = locator.bounding_box()
    if not box:
        locator.click(timeout=3_000)
        return

    x = box["x"] + max(8, box["width"] - 24)
    y = box["y"] + box["height"] / 2
    page.mouse.click(x, y)


def wait_for_any_text_visible(page: Page, texts: list[str], timeout_ms: int) -> bool:
    deadline = datetime.now().timestamp() + timeout_ms / 1000
    while datetime.now().timestamp() < deadline:
        for text in texts:
            try:
                if page.get_by_role("option", name=text, exact=True).first.is_visible():
                    return True
                if page.get_by_role("listbox").first.is_visible():
                    return True
                if page.get_by_text(text, exact=True).first.is_visible():
                    return True
                if page.get_by_text(text, exact=False).first.is_visible():
                    return True
            except Exception:
                continue
        page.wait_for_timeout(100)
    return False


def check_force_stop(force_stop_requested: StopFlag | None) -> None:
    if force_stop_requested and force_stop_requested():
        raise AutomationAborted("사용자가 강제종료를 요청했습니다.")


def click_dropdown_option(page: Page, candidates: list[str]) -> bool:
    last_error: Exception | None = None

    for candidate in candidates:
        try:
            option = page.get_by_role("option", name=candidate, exact=True).first
            option.wait_for(state="visible", timeout=5_000)
            option.scroll_into_view_if_needed()
            option.click()
            return True
        except PlaywrightTimeoutError as exc:
            last_error = exc

        try:
            option = page.get_by_role("option", name=re.compile(re.escape(candidate))).first
            option.wait_for(state="visible", timeout=5_000)
            option.scroll_into_view_if_needed()
            option.click()
            return True
        except PlaywrightTimeoutError as exc:
            last_error = exc

        try:
            option = page.get_by_text(candidate, exact=True).first
            option.wait_for(state="visible", timeout=5_000)
            option.scroll_into_view_if_needed()
            option.click()
            return True
        except PlaywrightTimeoutError as exc:
            last_error = exc

        try:
            option = page.get_by_text(candidate, exact=False).first
            option.wait_for(state="visible", timeout=5_000)
            option.scroll_into_view_if_needed()
            option.click()
            return True
        except PlaywrightTimeoutError as exc:
            last_error = exc

    return False


def select_dropdown_by_keyboard(page: Page, dropdown, target_value: str, field_info: dict[str, Any]) -> bool:
    options = field_info.get("options", [])
    if not options or target_value not in options:
        return False

    index = options.index(target_value)

    try:
        dropdown.focus()
        dropdown.press("Enter")
        if not wait_for_any_text_visible(page, [target_value], timeout_ms=1_000):
            dropdown.press("Space")

        for _ in range(index + 1):
            page.keyboard.press("ArrowDown")
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False


def dropdown_value_candidates(value: str, field_info: dict[str, Any]) -> list[str]:
    raw_value = str(value).strip()
    value_map = field_info.get("value_map", {})
    mapped_value = value_map.get(raw_value)
    if mapped_value is None:
        normalized_value_map = {str(key).replace(" ", ""): mapped for key, mapped in value_map.items()}
        mapped_value = normalized_value_map.get(raw_value.replace(" ", ""))

    candidates = [candidate for candidate in [mapped_value, raw_value] if candidate]
    return list(dict.fromkeys(candidates))


def click_export_experience(page: Page, value: str) -> None:
    """
    value가 O면 '수출 경험 있음' 클릭
    value가 X면 '처음입니다' 클릭
    """
    normalized = str(value).strip().upper()

    if normalized == "O":
        selector = SELECTORS["export_experience_has"]
    elif normalized == "X":
        selector = SELECTORS["export_experience_first"]
    else:
        raise ValueError(f"export_experience 값은 O 또는 X만 가능합니다. 현재 값: {value}")

    button = page.locator(selector).first
    button.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    button.scroll_into_view_if_needed()
    button.click()


def click_button_group_option(page: Page, field_info: dict[str, Any], value: str) -> None:
    value_map = field_info.get("value_map", {})
    label = value_map.get(value, value)
    option = page.get_by_text(str(label), exact=True).first
    option.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    option.scroll_into_view_if_needed()
    option.click()


def fill_excluded_countries(page: Page, row_data: dict[str, Any]) -> None:
    countries = split_country_values(row_data.get("excluded_countries", ""))
    if not countries:
        return

    input_locator = page.locator(SELECTORS["excluded_country_input"]).first
    input_locator.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)

    for country in countries:
        input_locator.scroll_into_view_if_needed()
        input_locator.fill(country)
        click_excluded_country_add_button(page, input_locator)
        page.wait_for_timeout(250)


def click_excluded_country_add_button(page: Page, input_locator) -> None:
    attempts = [
        lambda: input_locator.press("Enter", timeout=3_000),
        lambda: page.locator(SELECTORS["excluded_country_add_button"]).first.click(timeout=3_000),
        lambda: input_locator.locator("xpath=following-sibling::button[1]").click(timeout=3_000),
        lambda: input_locator.locator("xpath=../following-sibling::button[1]").click(timeout=3_000),
        lambda: click_right_of_locator(page, input_locator),
    ]

    last_error: Exception | None = None
    for attempt in attempts:
        try:
            attempt()
            return
        except Exception as exc:
            last_error = exc

    raise RuntimeError(f"분석 제외 국가 추가 버튼을 클릭하지 못했습니다: {last_error}")


def click_right_of_locator(page: Page, locator) -> None:
    box = locator.bounding_box()
    if not box:
        raise RuntimeError("입력칸 위치를 확인하지 못했습니다.")

    # 제외 국가 입력칸 오른쪽에 붙은 '+' 버튼 중앙을 좌표 fallback으로 클릭합니다.
    x = box["x"] + box["width"] + 34
    y = box["y"] + box["height"] / 2
    page.mouse.click(x, y)


def click_generate_button(page: Page) -> None:
    button = wait_for_first_visible(page, SELECTORS["generate_button"], ELEMENT_TIMEOUT_MS)
    button.scroll_into_view_if_needed()
    try:
        wait_until_enabled(page, button, ELEMENT_TIMEOUT_MS)
    except PlaywrightTimeoutError as exc:
        raise ValueError(
            "보고서 생성 버튼이 활성화되지 않습니다. "
            "필수 입력값(HS CODE, 수출액 규모, 수출 경험)이 누락되었거나 "
            "드롭다운 선택이 반영되지 않은 것 같습니다."
        ) from exc
    button.click()


def wait_until_enabled(page: Page, locator, timeout_ms: int) -> None:
    deadline = time.monotonic() + timeout_ms / 1000
    handle = locator.element_handle(timeout=timeout_ms)
    if handle is None:
        raise RuntimeError("버튼 요소를 찾지 못했습니다.")
    remaining_timeout_ms = max(1, int((deadline - time.monotonic()) * 1000))
    page.wait_for_function(
        "(el) => !el.disabled && el.getAttribute('aria-disabled') !== 'true'",
        arg=handle,
        timeout=remaining_timeout_ms,
    )


def download_button_selectors() -> list[str]:
    selectors = [SELECTORS["download_button"]]
    fallback_selector = SELECTORS.get("download_button_fallback")
    if fallback_selector:
        selectors.append(fallback_selector)
    return selectors


def wait_for_download_button(page: Page, timeout_ms: int = TIMEOUT_MS):
    deadline = datetime.now().timestamp() + timeout_ms / 1000
    candidates = download_button_selectors()

    while datetime.now().timestamp() < deadline:
        for selector in candidates:
            candidate = first_visible_locator(page, selector)
            if candidate is not None:
                return candidate
        page.wait_for_timeout(500)

    raise PlaywrightTimeoutError(f"다운로드 버튼이 {timeout_ms // 1000}초 안에 나타나지 않았습니다.")


def ensure_generation_failure_recorder(page: Page) -> list[tuple[float, str]]:
    """
    생성 스트리밍 API(POST)의 실패 이벤트를 페이지 수명 동안 상시 기록한다.

    대기 루프 진입 후에 리스너를 달면 '생성 클릭 직후~대기 시작 사이'에 끊긴
    요청을 놓치므로, 페이지 생성 시점에 한 번만 붙여서 (시각, 내용)으로 쌓아둔다.
    """
    recorder = getattr(page, "_generation_request_failures", None)
    if recorder is not None:
        return recorder

    recorder = []

    def on_request_failed(request) -> None:
        try:
            if request.method == "POST" and GENERATION_API_URL_MARKER in request.url:
                failure = getattr(request, "failure", None) or "원인 미상"
                recorder.append((time.time(), f"{request.method} {request.url} - {failure}"))
                if len(recorder) > 50:
                    del recorder[:25]
        except Exception:
            pass

    page.on("requestfailed", on_request_failed)
    setattr(page, "_generation_request_failures", recorder)
    return recorder


def page_text_signature(page: Page) -> str | None:
    """
    스톨 감지용 페이지 텍스트 지문. 실패하면 None(판정 보류).
    """
    try:
        return page.evaluate(
            "() => { const t = document.body ? document.body.textContent : '';"
            " let h = 0; for (let i = 0; i < t.length; i++) { h = (h * 31 + t.charCodeAt(i)) | 0; }"
            " return t.length + ':' + h; }"
        )
    except Exception:
        return None


def wait_for_download_or_generation_error(
    page: Page,
    timeout_ms: int = TIMEOUT_MS,
    status_callback: StatusCallback | None = None,
    force_stop_requested: StopFlag | None = None,
):
    deadline = datetime.now().timestamp() + timeout_ms / 1000
    started_at = datetime.now().timestamp()
    last_status_at = 0.0
    download_selectors = download_button_selectors()

    # 생성 스트리밍 POST가 끊기면(net::ERR_ABORTED 등) 화면만 봐서는 알 수 없으므로
    # 페이지 상시 레코더에서 이번 생성(클릭 직후 포함, 2초 여유) 이후의 실패만 본다.
    failure_recorder = ensure_generation_failure_recorder(page)
    failure_window_start = started_at - 2.0

    def recent_stream_failures() -> list[str]:
        return [description for ts, description in list(failure_recorder) if ts >= failure_window_start]

    stream_failed_at: float | None = None
    stall_signature: str | None = None
    stall_changed_at = started_at

    while datetime.now().timestamp() < deadline:
        check_force_stop(force_stop_requested)
        now = datetime.now().timestamp()
        if status_callback and now - last_status_at >= 10:
            elapsed = int(now - started_at)
            status_callback(f"보고서 생성 중입니다. 경과 {elapsed}초")
            last_status_at = now

        for selector in download_selectors:
            download_button = first_visible_locator(page, selector)
            if download_button is not None:
                if status_callback:
                    status_callback("PDF 저장 버튼이 나타났습니다. 다운로드를 시작합니다.")
                return "download", download_button
        retry_button = first_visible_locator(page, SELECTORS["retry_button"])
        if retry_button is not None or first_visible_locator(page, SELECTORS["streaming_error_text"]) is not None:
            if status_callback:
                status_callback("KOTRA 서버 오류 화면이 감지되었습니다.")
            return "error", retry_button
        if now - started_at > 5:
            new_analysis_button = first_visible_locator(page, SELECTORS["new_analysis_button"])
            if new_analysis_button is not None:
                if status_callback:
                    status_callback("분석이 중단되어 새로운 분석 시작 상태가 감지되었습니다.")
                return "analysis_stopped", new_analysis_button
            generate_button = first_visible_locator(page, SELECTORS["generate_button"])
            if generate_button is not None:
                if status_callback:
                    status_callback("초기 입력 화면으로 돌아온 상태가 감지되었습니다.")
                return "returned_to_form", generate_button

        stream_failures = recent_stream_failures()
        if stream_failures and stream_failed_at is None:
            stream_failed_at = now
            if status_callback:
                status_callback(
                    "보고서 생성 요청이 중단된 것이 감지되었습니다. "
                    f"{GENERATION_ABORT_GRACE_SECONDS}초 동안 화면 전환을 기다립니다."
                )
        if stream_failed_at is not None and now - stream_failed_at >= GENERATION_ABORT_GRACE_SECONDS:
            return "generation_aborted", "; ".join(stream_failures[-3:])

        signature = page_text_signature(page)
        if signature is None or signature != stall_signature:
            stall_signature = signature
            stall_changed_at = now
        elif now - stall_changed_at >= GENERATION_STALL_SECONDS:
            return "stalled", None

        page.wait_for_timeout(GENERATION_STATUS_POLL_INTERVAL_MS)

    raise PlaywrightTimeoutError(f"다운로드 버튼이 {timeout_ms // 1000}초 안에 나타나지 않았습니다.")


def is_visible(locator) -> bool:
    try:
        return locator.is_visible()
    except Exception:
        return False


def first_visible_locator(page: Page, selector: str, *, limit: int = 8):
    """
    selector와 일치하는 요소 중 '화면에 보이는' 첫 요소를 반환한다. 없으면 None.

    .first는 DOM 순서상 첫 요소에 고정되므로, 숨겨진 중복 노드(이전 화면 잔존
    요소, 반응형 레이아웃 중복 등)가 앞에 있으면 보이는 요소를 영영 못 찾는다.
    상태 감지는 반드시 이 함수로 '보이는 매치가 있는가'를 판정한다.
    """
    try:
        locator = page.locator(selector)
        count = min(locator.count(), limit)
    except Exception:
        return None
    for index in range(count):
        candidate = locator.nth(index)
        if is_visible(candidate):
            return candidate
    return None


def wait_for_first_visible(page: Page, selector: str, timeout_ms: int):
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        candidate = first_visible_locator(page, selector)
        if candidate is not None:
            return candidate
        if time.monotonic() >= deadline:
            raise PlaywrightTimeoutError(
                f"요소가 {timeout_ms // 1000}초 안에 화면에 나타나지 않았습니다: {selector}"
            )
        page.wait_for_timeout(250)


def setup_page_diagnostics(page: Page) -> list[str]:
    events: list[str] = []

    def remember(message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        events.append(f"[{timestamp}] {message}")
        if len(events) > 300:
            del events[:100]

    page.on("console", lambda msg: remember(f"console.{msg.type}: {msg.text}"))
    page.on("pageerror", lambda exc: remember(f"pageerror: {exc}"))
    page.on("requestfailed", lambda request: remember(f"requestfailed: {request.method} {request.url} - {request.failure}"))

    def on_response(response) -> None:
        if response.status >= 400:
            remember(f"response {response.status}: {response.url}")

    page.on("response", on_response)
    return events


def clear_diagnostics(events: list[str]) -> None:
    events.clear()


def save_failure_artifacts(
    page: Page,
    row_data: dict[str, Any],
    log_dir: str | Path,
    error_message: str,
    events: list[str],
    *,
    suffix: str = "",
) -> str:
    diagnostics_dir = Path(log_dir) / "diagnostics"
    diagnostics_dir.mkdir(parents=True, exist_ok=True)

    row_index = int(row_data.get("row_index", 0) or 0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix_part = f"_{safe_filename(suffix)}" if suffix else ""
    base_name = f"row_{row_index:03d}_{timestamp}{suffix_part}"
    screenshot_path = diagnostics_dir / f"{base_name}.png"
    text_path = diagnostics_dir / f"{base_name}.txt"

    try:
        page.screenshot(path=str(screenshot_path), full_page=True, timeout=SCREENSHOT_TIMEOUT_MS)
    except Exception as exc:
        events.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] full page screenshot failed: {exc}")
        try:
            # 전체 페이지 캡처가 폰트 로딩 등으로 매달리면 보이는 영역이라도 남긴다.
            page.screenshot(path=str(screenshot_path), full_page=False, timeout=SCREENSHOT_TIMEOUT_MS)
        except Exception as viewport_exc:
            events.append(
                f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] viewport screenshot failed: {viewport_exc}"
            )

    page_text = ""
    try:
        page_text = page.locator("body").inner_text(timeout=3_000)
    except Exception as exc:
        page_text = f"body text capture failed: {exc}"

    content = [
        f"row_index: {row_data.get('row_index', '')}",
        f"hs_code: {row_data.get('hs_code', '')}",
        f"product_name: {row_data.get('product_name', '')}",
        f"export_scale: {row_data.get('export_scale', '')}",
        f"export_experience: {row_data.get('export_experience', '')}",
        f"target_country: {row_data.get('target_country', '')}",
        f"excluded_countries: {row_data.get('excluded_countries', '')}",
        f"report_mode: {row_data.get('report_mode', '')}",
        f"url: {safe_page_url(page)}",
        f"error_message: {error_message}",
        "",
        "=== Recent Browser Events ===",
        *(events or ["(no captured browser events)"]),
        "",
        "=== Page Text Snapshot ===",
        page_text[:5000],
    ]
    text_path.write_text("\n".join(content), encoding="utf-8")
    return f"{text_path} / {screenshot_path}"


def safe_page_url(page: Page) -> str:
    try:
        return str(page.url)
    except Exception as exc:
        return f"(url unavailable: {exc})"


def is_closed_browser_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return (
        "target page, context or browser has been closed" in text
        or "browser has been closed" in text
        or "page has been closed" in text
        or "context has been closed" in text
    )


def summarize_country_failures(failures: list[tuple[str, Exception]]) -> RuntimeError:
    summary = "; ".join(f"{country}: {exc}" for country, exc in failures)
    return RuntimeError(f"일부 국가 보고서 생성에 실패했습니다: {summary}")


def download_report(
    page: Page,
    save_path: str | Path,
    timeout_ms: int = TIMEOUT_MS,
    status_callback: StatusCallback | None = None,
    force_stop_requested: StopFlag | None = None,
    *,
    row_data: dict[str, Any] | None = None,
    filename_pattern: str = "",
) -> Path:
    save_path = Path(save_path)
    if not (save_path.exists() and save_path.is_dir()):
        save_path.parent.mkdir(parents=True, exist_ok=True)

    status, payload = wait_for_download_or_generation_error(page, timeout_ms, status_callback, force_stop_requested)
    if status == "error":
        raise RuntimeError("KOTRA 보고서 생성 중 서버 스트리밍 오류가 발생했습니다.")
    if status == "returned_to_form":
        raise RuntimeError("보고서 생성 중 초기 입력 화면으로 돌아왔습니다. 서버 오류 또는 사용자의 되돌아가기 동작으로 판단됩니다.")
    if status == "analysis_stopped":
        raise RuntimeError("보고서 생성 중 분석이 중단되었습니다. 사용자가 분석 중단하기를 눌렀거나 사이트가 생성을 중단한 것으로 판단됩니다.")
    if status == "generation_aborted":
        detail = f" ({payload})" if payload else ""
        raise RuntimeError(
            f"보고서 생성 요청이 중단되었습니다. 서버가 생성을 취소했거나 네트워크가 끊긴 것으로 판단됩니다.{detail}"
        )
    if status == "stalled":
        raise RuntimeError(
            f"보고서 생성이 멈춘 것으로 판단됩니다. {GENERATION_STALL_SECONDS}초 동안 화면 변화가 없었습니다."
        )

    download_button = payload
    download_button.scroll_into_view_if_needed()
    wait_until_enabled(page, download_button, ELEMENT_TIMEOUT_MS)
    check_force_stop(force_stop_requested)

    with page.expect_download(timeout=timeout_ms) as download_info:
        download_button.click()

    download = download_info.value
    final_save_path = resolve_download_save_path(
        download,
        save_path,
        row_data=row_data,
        filename_pattern=filename_pattern,
    )
    download.save_as(str(final_save_path))
    if status_callback:
        status_callback("PDF 다운로드가 완료되었습니다.")
    return final_save_path


def resolve_download_save_path(
    download: Any,
    save_path: Path,
    *,
    row_data: dict[str, Any] | None = None,
    filename_pattern: str = "",
) -> Path:
    if not (save_path.exists() and save_path.is_dir()):
        return save_path

    suggested_filename = str(getattr(download, "suggested_filename", "") or "").strip() or "report.pdf"
    if normalize_filename_pattern(filename_pattern):
        filename = render_filename_pattern(filename_pattern, row_data or {}, suggested_filename=suggested_filename)
    else:
        filename = sanitize_report_filename(suggested_filename) or "report.pdf"
    return unique_file_path(save_path / filename)


def normalize_filename_pattern(filename_pattern: Any) -> str:
    return str(filename_pattern or "").strip()


def validate_filename_pattern(filename_pattern: Any) -> None:
    pattern = normalize_filename_pattern(filename_pattern)
    if not pattern:
        return

    remaining_text = FILENAME_PATTERN_TOKEN_RE.sub("", pattern)
    if "{" in remaining_text or "}" in remaining_text:
        raise ValueError("파일명 항목은 {hs_code}처럼 중괄호를 완성해서 입력해주세요.")

    unknown_tokens = sorted({token for token in FILENAME_PATTERN_TOKEN_RE.findall(pattern) if token not in FILENAME_PATTERN_TOKENS})
    if unknown_tokens:
        supported = ", ".join(f"{{{token}}}" for token in sorted(FILENAME_PATTERN_TOKENS))
        unknown = ", ".join(f"{{{token}}}" for token in unknown_tokens)
        raise ValueError(f"지원하지 않는 파일명 항목입니다: {unknown}. 지원 항목: {supported}")


def render_filename_pattern(
    filename_pattern: Any,
    row_data: dict[str, Any],
    *,
    suggested_filename: str = "report.pdf",
    now: datetime | None = None,
) -> str:
    pattern = normalize_filename_pattern(filename_pattern)
    if not pattern:
        return sanitize_report_filename(suggested_filename or "report.pdf") or "report.pdf"

    validate_filename_pattern(pattern)
    current_time = now or datetime.now()
    suggested_name = Path(str(suggested_filename or "report.pdf")).name
    suggested_path = Path(suggested_name)
    suffix = suggested_path.suffix or ".pdf"
    site_filename = suggested_path.stem if suggested_path.suffix else suggested_name

    values = {
        "row_index": str(row_data.get("row_index", "")).strip(),
        "company_name": str(row_data.get("company_name", "")).strip(),
        "business_number": str(row_data.get("business_number", "")).strip(),
        "hs_code": str(row_data.get("hs_code", "")).strip(),
        "product_name": str(row_data.get("product_name", "")).strip(),
        "target_country": str(row_data.get("target_country", "")).strip(),
        "excluded_countries": str(row_data.get("excluded_countries", "")).strip(),
        "report_mode": REPORT_MODE_LABELS.get(normalize_report_mode(row_data.get("report_mode", "")), str(row_data.get("report_mode", "")).strip()),
        "date": current_time.strftime("%Y%m%d"),
        "time": current_time.strftime("%H%M%S"),
        "datetime": current_time.strftime("%Y%m%d_%H%M%S"),
        "year": current_time.strftime("%Y"),
        "month": current_time.strftime("%m"),
        "day": current_time.strftime("%d"),
        "hour": current_time.strftime("%H"),
        "minute": current_time.strftime("%M"),
        "second": current_time.strftime("%S"),
        "site_filename": site_filename,
    }

    rendered = FILENAME_PATTERN_TOKEN_RE.sub(lambda match: values.get(match.group(1), ""), pattern)
    filename = sanitize_report_filename(rendered)
    if not filename:
        raise ValueError("파일명 패턴 결과가 비어 있습니다. 파일명 항목이나 문자를 추가해주세요.")

    if suffix and not filename.lower().endswith(suffix.lower()):
        filename = f"{filename}{suffix}"
    return filename


def sanitize_report_filename(filename: Any) -> str:
    safe = FILENAME_FORBIDDEN_CHARS_RE.sub("_", str(filename or ""))
    safe = FILENAME_CONTROL_CHARS_RE.sub("_", safe)
    safe = safe.strip().strip(".")
    if not safe:
        return ""

    suffix = Path(safe).suffix
    stem = Path(safe).stem if suffix else safe
    if stem.upper() in WINDOWS_RESERVED_FILENAMES:
        stem = f"_{stem}"

    max_stem_length = max(1, MAX_SAFE_FILENAME_LENGTH - len(suffix))
    if len(stem) > max_stem_length:
        stem = stem[:max_stem_length].rstrip(" .") or "report"

    return f"{stem}{suffix}"


def unique_file_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def reset_for_next_row(page: Page) -> None:
    # 빈 입력 화면이면(직전 단계에서 이미 초기화됨) 그대로 사용한다.
    # 값이 남아 있으면 이전 입력이 다음 보고서에 섞이지 않도록 초기화 경로를 끝까지 탄다.
    hs_code_input = first_visible_locator(page, SELECTORS["hs_code_input"])
    if hs_code_input is not None:
        try:
            if not hs_code_input.input_value(timeout=2_000).strip():
                return
        except Exception:
            pass

    new_analysis_button = first_visible_locator(page, SELECTORS["new_analysis_button"])
    if new_analysis_button is not None:
        new_analysis_button.scroll_into_view_if_needed()
        wait_until_enabled(page, new_analysis_button, ELEMENT_TIMEOUT_MS)
        new_analysis_button.click()
        wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)
        return

    try:
        reset_button = wait_for_first_visible(page, SELECTORS["reset_button"], 5_000)
        reset_button.scroll_into_view_if_needed()
        reset_button.click()
        wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        open_site(page)
        wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)


def ensure_initial_form(
    page: Page,
    status_callback: StatusCallback | None = None,
    wait_seconds: int = DEFAULT_INITIAL_FORM_WAIT_SECONDS,
) -> None:
    """
    저장된 브라우저 세션이 결과 화면을 복원해도 다음 자동화가 입력 화면에서 시작되도록 보장한다.
    """
    deadline = datetime.now().timestamp() + wait_seconds
    while datetime.now().timestamp() < deadline:
        if first_visible_locator(page, SELECTORS["hs_code_input"]) is not None:
            return
        new_analysis_button = first_visible_locator(page, SELECTORS["new_analysis_button"])
        if new_analysis_button is not None:
            if status_callback:
                status_callback("이전 결과 화면이 감지되어 새 분석 화면으로 돌아갑니다.")
            new_analysis_button.scroll_into_view_if_needed()
            wait_until_enabled(page, new_analysis_button, ELEMENT_TIMEOUT_MS)
            new_analysis_button.click()
            wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)
            return
        reset_button = first_visible_locator(page, SELECTORS["reset_button"])
        if reset_button is not None:
            if status_callback:
                status_callback("입력 화면을 초기화합니다.")
            reset_button.scroll_into_view_if_needed()
            reset_button.click()
            wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)
            return
        page.wait_for_timeout(300)

    if status_callback:
        status_callback("초기 입력 화면이 보이지 않아 KOTRA 페이지를 새로 엽니다.")
    open_site(page)
    wait_for_first_visible(page, SELECTORS["hs_code_input"], ELEMENT_TIMEOUT_MS)


def retry_generation(page: Page) -> None:
    retry_button = wait_for_first_visible(page, SELECTORS["retry_button"], ELEMENT_TIMEOUT_MS)
    retry_button.scroll_into_view_if_needed()
    wait_until_enabled(page, retry_button, ELEMENT_TIMEOUT_MS)
    retry_button.click()


def page_wait_seconds(seconds: int, force_stop_requested: StopFlag | None = None) -> None:
    deadline = datetime.now().timestamp() + seconds
    while datetime.now().timestamp() < deadline:
        check_force_stop(force_stop_requested)
        remaining = deadline - datetime.now().timestamp()
        time.sleep(max(0.1, min(0.5, remaining)))


def process_row(
    page: Page,
    row_data: dict[str, Any],
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int = TIMEOUT_MS,
    retry_count: int = GENERATION_RETRY_COUNT,
    status_callback: StatusCallback | None = None,
    force_stop_requested: StopFlag | None = None,
    filename_pattern: str = "",
    direct_report_count: int = DEFAULT_DIRECT_REPORT_COUNT,
    task_status_callback: TaskStatusCallback | None = None,
    defer_country_failures_for_retry: bool = False,
) -> RowProcessResult:
    report_mode = normalize_report_mode(row_data.get("report_mode", REPORT_MODE_DIRECT))
    direct_report_count = normalize_direct_report_count(row_data.get("direct_report_count", direct_report_count))
    if report_mode == REPORT_MODE_RECOMMEND:
        return process_recommendation_row(
            page,
            row_data,
            save_path,
            log_dir,
            diagnostic_events,
            timeout_ms,
            retry_count,
            status_callback,
            force_stop_requested,
            filename_pattern,
            direct_report_count,
            task_status_callback,
            defer_country_failures_for_retry,
        )

    return process_direct_row(
        page,
        row_data,
        save_path,
        log_dir,
        diagnostic_events,
        timeout_ms,
        retry_count,
        status_callback,
        force_stop_requested,
        filename_pattern,
        direct_report_count,
        task_status_callback,
        defer_country_failures_for_retry,
    )


def process_direct_row(
    page: Page,
    row_data: dict[str, Any],
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int,
    retry_count: int,
    status_callback: StatusCallback | None,
    force_stop_requested: StopFlag | None,
    filename_pattern: str,
    direct_report_count: int,
    task_status_callback: TaskStatusCallback | None = None,
    defer_country_failures_for_retry: bool = False,
) -> RowProcessResult:
    direct_report_count = normalize_direct_report_count(direct_report_count)
    target_countries = split_country_values(row_data.get("target_country", ""))[:direct_report_count]
    if not target_countries:
        raise ValueError("희망 진출 국가가 없어 수출시장 분석 보고서를 생성할 수 없습니다.")

    def emit_task_status(country: str, status: str, saved_files: list[Path] | None = None) -> None:
        if task_status_callback:
            row_index = int(row_data.get("row_index", 0))
            task_status_callback({
                "ui_key": str(row_data.get("ui_key", "") or f"row:{row_index}"),
                "row_index": row_index,
                "task_type": TASK_TYPE_DIRECT,
                "country": country,
                "status": status,
                "saved_files": [str(path) for path in (saved_files or [])],
                "ts": time.time(),
            })

    for country in target_countries:
        emit_task_status(country, STATUS_PENDING)

    direct_report_files: list[Path] = []
    country_failures: list[tuple[str, Exception]] = []
    for index, country in enumerate(target_countries):
        completed_task = completed_report_task(log_dir, row_data, TASK_TYPE_DIRECT, country)
        if completed_task is not None:
            saved_file = Path(str(completed_task.get("saved_file", "")))
            direct_report_files.append(saved_file)
            row_data["direct_report_files"] = join_path_values(direct_report_files)
            if status_callback:
                status_callback(f"이미 완료된 수출시장 분석 보고서를 건너뜁니다: {country}")
            emit_task_status(country, STATUS_SUCCESS, [saved_file])
            continue

        update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_RUNNING)
        emit_task_status(country, STATUS_RUNNING)
        try:
            if index > 0:
                reset_for_next_row(page)
            saved_direct_report = process_direct_country_report(
                page,
                row_data,
                country,
                save_path,
                log_dir,
                diagnostic_events,
                timeout_ms,
                retry_count,
                status_callback,
                force_stop_requested,
                filename_pattern,
            )
        except Exception as exc:
            if page_is_closed(page) or is_closed_browser_error(exc):
                update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_FAILED, error_message=str(exc))
                emit_task_status(country, STATUS_FAILED)
                raise
            failed_status = STATUS_RETRY_PENDING if defer_country_failures_for_retry else STATUS_FAILED
            update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, failed_status, error_message=str(exc))
            emit_task_status(country, failed_status)
            country_failures.append((country, exc))
            if status_callback:
                next_action = "남은 국가 처리 후 자동 재시도합니다" if defer_country_failures_for_retry else "남은 국가를 계속 처리합니다"
                status_callback(f"{country} 보고서 생성 실패, {next_action}: {exc}")
            continue
        direct_report_files.append(saved_direct_report)
        row_data["direct_report_files"] = join_path_values(direct_report_files)
        update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_SUCCESS, saved_file=saved_direct_report)
        emit_task_status(country, STATUS_SUCCESS, [saved_direct_report])

    row_data["final_target_countries"] = join_country_values(target_countries)
    if country_failures:
        raise summarize_country_failures(country_failures)
    return RowProcessResult(
        saved_files=direct_report_files,
        direct_report_files=direct_report_files,
        final_target_countries=target_countries,
    )


def process_recommendation_row(
    page: Page,
    row_data: dict[str, Any],
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int,
    retry_count: int,
    status_callback: StatusCallback | None,
    force_stop_requested: StopFlag | None,
    filename_pattern: str,
    direct_report_count: int,
    task_status_callback: TaskStatusCallback | None = None,
    defer_country_failures_for_retry: bool = False,
) -> RowProcessResult:
    direct_report_count = normalize_direct_report_count(direct_report_count)
    recommend_then_direct = truthy(row_data.get("recommend_then_direct", False))

    def emit_task_status(task_type: str, country: str, status: str, saved_files: list[Path] | None = None) -> None:
        if task_status_callback and recommend_then_direct:
            row_index = int(row_data.get("row_index", 0))
            task_status_callback({
                "ui_key": str(row_data.get("ui_key", "") or f"row:{row_index}"),
                "row_index": row_index,
                "task_type": task_type,
                "country": country,
                "status": status,
                "saved_files": [str(path) for path in (saved_files or [])],
                "ts": time.time(),
            })

    completed_recommend_task = completed_report_task(log_dir, row_data, TASK_TYPE_RECOMMEND)
    completed_final_targets = (
        split_country_values(completed_recommend_task.get("final_target_countries", ""))
        if completed_recommend_task is not None
        else []
    )
    can_skip_recommend = completed_recommend_task is not None and (
        not recommend_then_direct or len(completed_final_targets) >= direct_report_count
    )

    if can_skip_recommend:
        recommendation_report = Path(str(completed_recommend_task.get("saved_file", "")))
        row_data["recommendation_report_file"] = str(recommendation_report)
        row_data["recommended_countries"] = str(completed_recommend_task.get("recommended_countries", ""))
        row_data["final_target_countries"] = str(completed_recommend_task.get("final_target_countries", ""))
        if status_callback:
            status_callback("이미 완료된 유망 시장 추천 보고서를 건너뜁니다.")
        emit_task_status(TASK_TYPE_RECOMMEND, "", STATUS_SUCCESS, [recommendation_report])
    else:
        update_report_task_status(log_dir, row_data, TASK_TYPE_RECOMMEND, "", STATUS_RUNNING)
        emit_task_status(TASK_TYPE_RECOMMEND, "", STATUS_RUNNING)
        try:
            recommendation_report = process_recommendation_report(
                page,
                row_data,
                save_path,
                log_dir,
                diagnostic_events,
                timeout_ms,
                retry_count,
                status_callback,
                force_stop_requested,
                filename_pattern,
            )
            row_data["recommendation_report_file"] = str(recommendation_report)
        except Exception as exc:
            update_report_task_status(log_dir, row_data, TASK_TYPE_RECOMMEND, "", STATUS_FAILED, error_message=str(exc))
            emit_task_status(TASK_TYPE_RECOMMEND, "", STATUS_FAILED)
            raise

    if not recommend_then_direct:
        if not can_skip_recommend:
            update_report_task_status(log_dir, row_data, TASK_TYPE_RECOMMEND, "", STATUS_SUCCESS, saved_file=recommendation_report)
        return RowProcessResult(
            saved_files=[recommendation_report],
            recommendation_report_file=recommendation_report,
        )

    if can_skip_recommend:
        recommended = split_country_values(row_data.get("recommended_countries", ""))
        final_targets = split_country_values(row_data.get("final_target_countries", ""))[:direct_report_count]
    else:
        try:
            existing_targets = split_country_values(row_data.get("target_country", ""))[:direct_report_count]
            needed_count = direct_report_count - len(existing_targets)
            recommended = extract_recommended_countries(page, row_data, needed_count)
            final_targets = (existing_targets + recommended)[:direct_report_count]
            row_data["recommended_countries"] = join_country_values(recommended)
            row_data["final_target_countries"] = join_country_values(final_targets)

            if len(final_targets) < direct_report_count:
                raise ValueError(f"추천 결과에서 직접 분석 대상 국가 {direct_report_count}개를 확보하지 못했습니다.")
            update_report_task_status(log_dir, row_data, TASK_TYPE_RECOMMEND, "", STATUS_SUCCESS, saved_file=recommendation_report)
            emit_task_status(TASK_TYPE_RECOMMEND, "", STATUS_SUCCESS, [recommendation_report])
        except Exception as exc:
            update_report_task_status(log_dir, row_data, TASK_TYPE_RECOMMEND, "", STATUS_FAILED, saved_file=recommendation_report, error_message=str(exc))
            emit_task_status(TASK_TYPE_RECOMMEND, "", STATUS_FAILED)
            raise

    for country in final_targets:
        emit_task_status(TASK_TYPE_DIRECT, country, STATUS_PENDING)

    direct_report_files: list[Path] = []
    country_failures: list[tuple[str, Exception]] = []
    for country in final_targets:
        completed_direct_task = completed_report_task(log_dir, row_data, TASK_TYPE_DIRECT, country)
        if completed_direct_task is not None:
            saved_file = Path(str(completed_direct_task.get("saved_file", "")))
            direct_report_files.append(saved_file)
            row_data["direct_report_files"] = join_path_values(direct_report_files)
            if status_callback:
                status_callback(f"이미 완료된 수출시장 분석 보고서를 건너뜁니다: {country}")
            emit_task_status(TASK_TYPE_DIRECT, country, STATUS_SUCCESS, [saved_file])
            continue

        update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_RUNNING)
        emit_task_status(TASK_TYPE_DIRECT, country, STATUS_RUNNING)
        try:
            reset_for_next_row(page)
            saved_direct_report = process_direct_country_report(
                page,
                row_data,
                country,
                save_path,
                log_dir,
                diagnostic_events,
                timeout_ms,
                retry_count,
                status_callback,
                force_stop_requested,
                filename_pattern,
            )
        except Exception as exc:
            if page_is_closed(page) or is_closed_browser_error(exc):
                update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_FAILED, error_message=str(exc))
                emit_task_status(TASK_TYPE_DIRECT, country, STATUS_FAILED)
                raise
            failed_status = STATUS_RETRY_PENDING if defer_country_failures_for_retry else STATUS_FAILED
            update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, failed_status, error_message=str(exc))
            emit_task_status(TASK_TYPE_DIRECT, country, failed_status)
            country_failures.append((country, exc))
            if status_callback:
                next_action = "남은 국가 처리 후 자동 재시도합니다" if defer_country_failures_for_retry else "남은 국가를 계속 처리합니다"
                status_callback(f"{country} 보고서 생성 실패, {next_action}: {exc}")
            continue
        direct_report_files.append(saved_direct_report)
        row_data["direct_report_files"] = join_path_values(direct_report_files)
        update_report_task_status(log_dir, row_data, TASK_TYPE_DIRECT, country, STATUS_SUCCESS, saved_file=saved_direct_report)
        emit_task_status(TASK_TYPE_DIRECT, country, STATUS_SUCCESS, [saved_direct_report])

    if country_failures:
        raise summarize_country_failures(country_failures)
    return RowProcessResult(
        saved_files=[recommendation_report, *direct_report_files],
        recommendation_report_file=recommendation_report,
        direct_report_files=direct_report_files,
        recommended_countries=recommended,
        final_target_countries=final_targets,
    )


def process_direct_country_report(
    page: Page,
    row_data: dict[str, Any],
    country: str,
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int,
    retry_count: int,
    status_callback: StatusCallback | None,
    force_stop_requested: StopFlag | None,
    filename_pattern: str,
) -> Path:
    direct_row = {**row_data, "target_country": country}
    check_force_stop(force_stop_requested)
    select_direct_country_analysis(page)
    check_force_stop(force_stop_requested)
    fill_form(page, direct_row, DIRECT_FIELD_MAPPING)
    check_force_stop(force_stop_requested)
    click_generate_button(page)
    return submit_and_download_report(
        page,
        direct_row,
        save_path,
        log_dir,
        diagnostic_events,
        timeout_ms,
        retry_count,
        status_callback,
        force_stop_requested,
        filename_pattern,
    )


def process_recommendation_report(
    page: Page,
    row_data: dict[str, Any],
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int,
    retry_count: int,
    status_callback: StatusCallback | None,
    force_stop_requested: StopFlag | None,
    filename_pattern: str,
) -> Path:
    check_force_stop(force_stop_requested)
    select_recommend_market_analysis(page)
    check_force_stop(force_stop_requested)
    fill_form(page, row_data, RECOMMEND_FIELD_MAPPING)
    fill_excluded_countries(page, row_data)
    check_force_stop(force_stop_requested)
    click_generate_button(page)
    return submit_and_download_report(
        page,
        row_data,
        save_path,
        log_dir,
        diagnostic_events,
        timeout_ms,
        retry_count,
        status_callback,
        force_stop_requested,
        filename_pattern,
    )


def submit_and_download_report(
    page: Page,
    row_data: dict[str, Any],
    save_path: str | Path,
    log_dir: str | Path,
    diagnostic_events: list[str],
    timeout_ms: int,
    retry_count: int,
    status_callback: StatusCallback | None,
    force_stop_requested: StopFlag | None,
    filename_pattern: str,
) -> Path:
    last_error: Exception | None = None
    for attempt in range(retry_count + 1):
        try:
            if status_callback:
                status_callback("보고서 생성 요청을 보냈습니다.")
            return download_report(
                page,
                save_path,
                timeout_ms,
                status_callback,
                force_stop_requested,
                row_data=row_data,
                filename_pattern=filename_pattern,
            )
        except RuntimeError as exc:
            last_error = exc
            message = str(exc)
            # 페이지를 더 조작하기 전에(화면이 바뀌기 전에) 실패 순간의 진단을 남긴다.
            fatal_markers = (
                ("초기 입력 화면", "returned_to_form"),
                ("생성 요청이 중단", "generation_aborted"),
                ("생성이 멈춘", "generation_stalled"),
                ("분석이 중단", "analysis_stopped"),
            )
            fatal_suffix = next((suffix for marker, suffix in fatal_markers if marker in message), None)
            if fatal_suffix:
                artifacts = save_failure_artifacts(
                    page,
                    row_data,
                    log_dir,
                    message,
                    diagnostic_events,
                    suffix=fatal_suffix,
                )
                raise GenerationError(message, artifacts) from exc

            if "서버 스트리밍 오류" not in message or attempt >= retry_count:
                raise

            artifacts = save_failure_artifacts(
                page,
                row_data,
                log_dir,
                str(exc),
                diagnostic_events,
                suffix=f"streaming_error_attempt_{attempt + 1}",
            )
            if status_callback:
                status_callback(f"KOTRA 서버 스트리밍 오류 감지: {attempt + 1}/{retry_count}회 재시도합니다. 진단 저장: {artifacts}")
            check_force_stop(force_stop_requested)
            retry_generation(page)

    raise RuntimeError("보고서 생성 재시도 후에도 실패했습니다.") from last_error


def normalize_report_mode(value: Any) -> str:
    text = str(value or REPORT_MODE_DIRECT).strip().lower()
    if text in {REPORT_MODE_RECOMMEND, "recommendation", "유망 시장 추천 보고서 생성", "유망 시장 추천", "추천"}:
        return REPORT_MODE_RECOMMEND
    return REPORT_MODE_DIRECT


def truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"1", "true", "yes", "y", "o", "on", "사용", "사용함", "켜짐"}


# 대상국/제외국가 칸에 '-' 처럼 대시만 적은 값은 '희망 국가 없음' 의도의 플레이스홀더로 본다.
COUNTRY_NONE_PLACEHOLDER_RE = re.compile(r"[-–—―－]+")


def split_country_values(value: Any) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []
    parts = re.split(r"[,/\n]+", text)
    countries: list[str] = []
    seen: set[str] = set()
    for part in parts:
        country = normalize_single_country(part)
        key = normalize_country_key(country)
        if country and key not in seen:
            seen.add(key)
            countries.append(country)
    return countries


def normalize_single_country(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)
    text = text.strip(" ,")
    if COUNTRY_NONE_PLACEHOLDER_RE.fullmatch(text):
        return ""
    return text


def join_country_values(countries: list[str]) -> str:
    return ", ".join(country for country in countries if country)


def normalize_country_key(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").strip()).lower()


def extract_recommended_countries(page: Page, row_data: dict[str, Any], needed_count: int) -> list[str]:
    if needed_count <= 0:
        return []

    blocked = {
        normalize_country_key(country)
        for country in [
            *split_country_values(row_data.get("target_country", "")),
            *split_country_values(row_data.get("excluded_countries", "")),
        ]
    }
    selected: list[str] = []
    selected_keys: set[str] = set()

    for country in recommended_country_candidates_by_card(page):
        key = normalize_country_key(country)
        if not key or key in blocked or key in selected_keys:
            continue
        selected.append(country)
        selected_keys.add(key)
        if len(selected) >= needed_count:
            return selected

    for country in recommended_country_candidates_from_section(page):
        key = normalize_country_key(country)
        if not key or key in blocked or key in selected_keys:
            continue
        selected.append(country)
        selected_keys.add(key)
        if len(selected) >= needed_count:
            return selected

    return selected


def recommended_country_candidates_by_card(page: Page) -> list[str]:
    candidates: list[str] = []
    try:
        cards = page.locator(SELECTORS["market_analysis_card"])
        for index in range(cards.count()):
            chips = cards.nth(index).locator(SELECTORS["market_country_chip"])
            if chips.count() == 0:
                continue
            text = normalize_recommended_country_text(chips.first.inner_text(timeout=2_000))
            if text:
                candidates.append(text)
    except Exception:
        return candidates
    return candidates


def recommended_country_candidates_from_section(page: Page) -> list[str]:
    try:
        section = page.locator(SELECTORS["market_analysis_section"]).first
        section.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        texts = section.locator("button").all_inner_texts()
    except Exception:
        return []

    candidates: list[str] = []
    for text in texts:
        country = normalize_recommended_country_text(text)
        if country:
            candidates.append(country)
    return candidates


def normalize_recommended_country_text(value: Any) -> str:
    text = normalize_single_country(value)
    if not text or len(text) > 30:
        return ""
    if any(token in text for token in ("PDF", "저장", "새로운 분석", "보고서")):
        return ""
    return text


def normalize_parallel_sessions(parallel_sessions: int) -> int:
    try:
        count = int(parallel_sessions)
    except (TypeError, ValueError):
        return 1
    return max(1, min(MAX_PARALLEL_SESSIONS, count))


def normalize_direct_report_count(direct_report_count: int) -> int:
    try:
        count = int(direct_report_count)
    except (TypeError, ValueError):
        return DEFAULT_DIRECT_REPORT_COUNT
    return max(1, min(MAX_DIRECT_REPORT_COUNT, count))


def normalize_row_retry_count(row_retry_count: int) -> int:
    try:
        count = int(row_retry_count)
    except (TypeError, ValueError):
        return DEFAULT_ROW_RETRY_COUNT
    return max(0, count)


def build_failure_error_message(
    page: Page,
    row_data: dict[str, Any],
    log_dir: str | Path,
    diagnostic_events: list[str],
    exc: Exception,
    *,
    suffix: str = "",
) -> tuple[str, str]:
    if isinstance(exc, PlaywrightTimeoutError):
        message = f"제한 시간 안에 필요한 요소를 찾지 못했습니다: {exc}"
        artifacts = save_failure_artifacts(page, row_data, log_dir, message, diagnostic_events, suffix=suffix)
        return message, f"{message} / diagnostics: {artifacts}"

    if isinstance(exc, GenerationError):
        error_message = f"{exc} / diagnostics: {exc.artifacts}" if exc.artifacts else str(exc)
        return str(exc), error_message

    message = str(exc)
    artifacts = save_failure_artifacts(page, row_data, log_dir, message, diagnostic_events, suffix=suffix)
    return message, f"{message} / diagnostics: {artifacts}"


def retry_suffix(attempt: int, final: bool = False) -> str:
    label = "final" if final else "retry"
    return f"{label}_attempt_{attempt + 1}"


def is_retry_enabled(auto_retry_enabled: RetryFlag | None) -> bool:
    if auto_retry_enabled is None:
        return True
    try:
        return bool(auto_retry_enabled())
    except Exception:
        return True


def prepare_parallel_storage_state(
    state_path: Path,
    headless: bool,
    emit_status: StatusCallback,
) -> bool:
    if headless:
        emit_status("백그라운드 실행에서는 수동 로그인 대기를 건너뜁니다.")
        return False

    with sync_playwright() as playwright:
        browser = launch_browser(playwright, headless, emit_status)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_navigation_timeout(PAGE_LOAD_TIMEOUT_MS)
        page.set_default_timeout(ELEMENT_TIMEOUT_MS)
        try:
            emit_status("병렬 실행 전 로그인 세션을 준비합니다.")
            open_site(page)
            ensure_initial_form(page, emit_status)
            emit_status("로그인이 필요하면 브라우저에서 직접 로그인한 뒤 Enter를 누르세요.")
            input("로그인 완료 후 Enter를 누르세요: ")
            ensure_initial_form(page, emit_status)
            context.storage_state(path=str(state_path))
            emit_status(f"로그인 세션을 저장했습니다: {state_path}")
            return True
        finally:
            context.close()
            browser.close()


def run_parallel_automation(
    input_excel_path: Path,
    download_dir: Path,
    headless: bool,
    *,
    log_dir: Path,
    state_path: Path,
    rows: list[dict[str, Any]],
    timeout_ms: int,
    retry_count: int,
    use_storage_state: bool,
    save_storage_state: bool,
    retry_failed_only: bool,
    wait_for_manual_login: bool,
    parallel_sessions: int,
    row_retry_count: int,
    direct_report_count: int,
    auto_retry_enabled: RetryFlag | None,
    filename_pattern: str,
    status_callback: StatusCallback | None,
    progress_callback: ProgressCallback | None,
    stop_requested: StopFlag | None,
    force_stop_requested: StopFlag | None,
    row_status_callback: RowStatusCallback | None = None,
    task_status_callback: TaskStatusCallback | None = None,
) -> dict[str, Any]:
    total = len(rows)
    worker_count = min(normalize_parallel_sessions(parallel_sessions), total)
    row_retry_count = normalize_row_retry_count(row_retry_count)
    direct_report_count = normalize_direct_report_count(direct_report_count)
    row_queue: queue.Queue[tuple[int, dict[str, Any], int]] = queue.Queue()
    for row_number, row_data in enumerate(rows, start=1):
        row_queue.put((row_number, row_data, 0))

    counter_lock = threading.Lock()
    file_lock = threading.Lock()
    storage_state_lock = threading.Lock()
    wait_status_lock = threading.Lock()
    wait_statuses: dict[int, tuple[str, int]] = {}
    storage_state_saved = False
    active_workers = worker_count
    success_count = 0
    failed_count = 0
    completed_count = 0
    completed_row_numbers: set[int] = set()
    retry_pending_failed_numbers: set[int] = set()
    last_wait_summary_at = 0.0

    def emit_status(message: str) -> None:
        if status_callback:
            status_callback(message)
        else:
            print(message)

    def emit_progress(status: str = "") -> None:
        if not progress_callback:
            return
        with counter_lock:
            progress_callback(
                {
                    "total": total,
                    "current": completed_count,
                    "success": success_count,
                    "failed": failed_count,
                    "status": status,
                }
            )

    def mark_result(current_index: int, success: bool) -> None:
        nonlocal success_count, failed_count, completed_count
        with counter_lock:
            was_retry_pending_failure = current_index in retry_pending_failed_numbers
            if was_retry_pending_failure:
                retry_pending_failed_numbers.remove(current_index)
            if success:
                if was_retry_pending_failure:
                    failed_count = max(0, failed_count - 1)
                success_count += 1
            elif not was_retry_pending_failure:
                failed_count += 1
            completed_count += 1
            completed_row_numbers.add(current_index)

    def mark_retry_pending_failure(current_index: int) -> None:
        nonlocal failed_count
        with counter_lock:
            if current_index not in retry_pending_failed_numbers:
                retry_pending_failed_numbers.add(current_index)
                failed_count += 1

    def clear_wait_status(session_id: int) -> None:
        with wait_status_lock:
            wait_statuses.pop(session_id, None)

    def emit_wait_summary(session_id: int, row_label: str, elapsed: int) -> None:
        nonlocal last_wait_summary_at
        now = datetime.now().timestamp()
        with wait_status_lock:
            wait_statuses[session_id] = (row_label, elapsed)
            if now - last_wait_summary_at < PARALLEL_WAIT_SUMMARY_INTERVAL_SECONDS:
                return
            last_wait_summary_at = now
            parts = [
                f"세션 {item_session_id}({item_row_label}, {item_elapsed}초)"
                for item_session_id, (item_row_label, item_elapsed) in sorted(wait_statuses.items())
            ]
        emit_progress("생성 대기 중: " + ", ".join(parts))

    def emit_row_progress(session_id: int, row_label: str, prefix: str, message: str) -> None:
        match = re.search(r"보고서 생성 중입니다\. 경과 (\d+)초", message)
        if match:
            emit_wait_summary(session_id, row_label, int(match.group(1)))
            return

        if (
            "PDF 저장 버튼" in message
            or "PDF 다운로드" in message
            or "오류 화면" in message
            or "초기 입력 화면" in message
            or "분석이 중단" in message
            or "생성 요청이 중단" in message
            or "생성이 멈춘" in message
        ):
            clear_wait_status(session_id)
        emit_progress(f"{prefix} {message}")

    def combined_force_stop_requested() -> bool:
        return bool(force_stop_requested and force_stop_requested())

    def should_stop_before_next_row() -> bool:
        return bool(stop_requested and stop_requested()) or combined_force_stop_requested()

    if wait_for_manual_login:
        prepared = prepare_parallel_storage_state(state_path, headless, emit_status)
        use_storage_state = use_storage_state or prepared

    emit_status(f"병렬 처리 모드로 실행합니다: {worker_count}개 세션")
    emit_progress("병렬 처리 준비 중")

    def finish_worker(context: Any, session_id: int) -> None:
        nonlocal active_workers, storage_state_saved
        with storage_state_lock:
            active_workers -= 1
            should_save = save_storage_state and not storage_state_saved and active_workers == 0
            if should_save:
                try:
                    context.storage_state(path=str(state_path))
                    storage_state_saved = True
                    emit_status(f"[세션 {session_id}] 브라우저 세션을 저장했습니다: {state_path}")
                except Exception as exc:
                    emit_status(f"[세션 {session_id}] 브라우저 세션 저장을 건너뜁니다: {exc}")

    def finish_worker_without_context() -> None:
        nonlocal active_workers
        with storage_state_lock:
            active_workers -= 1

    def emit_row_status_for(
        session_id: int,
        row_data: dict[str, Any],
        status: str,
        saved_files: list[Path] | None = None,
    ) -> None:
        if row_status_callback:
            row_index = int(row_data.get("row_index", 0))
            row_status_callback({
                "ui_key": str(row_data.get("ui_key", "") or f"row:{row_index}"),
                "row_index": row_index,
                "status": status,
                "session_id": session_id,
                "product_name": str(row_data.get("product_name", "")),
                "hs_code": str(row_data.get("hs_code", "")),
                "target_country": str(row_data.get("target_country", "")),
                "report_mode": str(row_data.get("report_mode", "")),
                "recommend_then_direct": bool(row_data.get("recommend_then_direct", False)),
                "saved_files": [str(path) for path in (saved_files or [])],
                "ts": time.time(),
            })

    def run_worker(session_id: int) -> None:
        def emit_row_status(row_data: dict[str, Any], status: str, saved_files: list[Path] | None = None) -> None:
            emit_row_status_for(session_id, row_data, status, saved_files)

        with sync_playwright() as playwright:
            start_delay = (session_id - 1) * PARALLEL_SESSION_START_DELAY_SECONDS
            if start_delay:
                emit_status(f"[세션 {session_id}] 서버 부하 분산을 위해 {start_delay}초 후 시작합니다.")
                page_wait_seconds(start_delay, force_stop_requested)

            try:
                browser = launch_browser(playwright, headless, lambda message: emit_status(f"[세션 {session_id}] {message}"))
            except RuntimeError as exc:
                finish_worker_without_context()
                with file_lock:
                    error_path = write_startup_error(log_dir, f"[세션 {session_id}] {exc}")
                emit_status(f"[세션 {session_id}] 브라우저 시작 오류를 기록했습니다: {error_path}")
                raise

            context, page, diagnostic_events = create_context_page(
                browser,
                use_storage_state=use_storage_state,
                state_path=state_path,
            )

            def recover_worker_page(reason: str) -> None:
                nonlocal browser, context, page, diagnostic_events
                emit_status(f"[세션 {session_id}] 브라우저 세션을 복구합니다: {reason}")
                close_browser_context(context, None)
                if browser_is_closed(browser):
                    close_browser_context(None, browser)
                    browser = launch_browser(
                        playwright,
                        headless,
                        lambda message: emit_status(f"[세션 {session_id}] {message}"),
                    )
                context, page, diagnostic_events = create_context_page(
                    browser,
                    use_storage_state=use_storage_state,
                    state_path=state_path,
                )
                open_site(page)
                ensure_initial_form(
                    page,
                    lambda message: emit_status(f"[세션 {session_id}] {message}"),
                    wait_seconds=PARALLEL_INITIAL_FORM_WAIT_SECONDS,
                )

            try:
                emit_status(f"[세션 {session_id}] KOTRA 보고서 생성 페이지에 접속합니다.")
                open_site(page)
                ensure_initial_form(
                    page,
                    lambda message: emit_status(f"[세션 {session_id}] {message}"),
                    wait_seconds=PARALLEL_INITIAL_FORM_WAIT_SECONDS,
                )

                next_retry_item: tuple[int, dict[str, Any], int] | None = None

                while True:
                    if combined_force_stop_requested():
                        emit_status(f"[세션 {session_id}] 강제종료 요청으로 작업을 즉시 중단합니다.")
                        break
                    if should_stop_before_next_row():
                        emit_status(f"[세션 {session_id}] 중지 요청이 있어 다음 행으로 넘어가지 않습니다.")
                        break

                    from_queue = False
                    if next_retry_item is not None:
                        current_index, row_data, row_attempt = next_retry_item
                        next_retry_item = None
                    else:
                        try:
                            current_index, row_data, row_attempt = row_queue.get_nowait()
                            from_queue = True
                        except queue.Empty:
                            break

                    row_index = int(row_data["row_index"])
                    row_label = f"{current_index}/{total}"
                    if retry_failed_only:
                        row_label = f"{row_label}, 원본 행 {row_index}"
                    if row_attempt:
                        row_label = f"{row_label}, 자동 재시도 {row_attempt}/{row_retry_count}"
                    prefix = f"[세션 {session_id}] [{row_label}]"
                    should_requeue = False

                    try:
                        emit_progress(f"{prefix} 입력 및 보고서 생성 중")
                        emit_status(f"{prefix} 행 처리를 시작합니다.")
                        clear_diagnostics(diagnostic_events)
                        row_data["use_task_resume"] = bool(retry_failed_only or row_attempt > 0)
                        with file_lock:
                            update_processing_status(input_excel_path, log_dir, row_data, STATUS_RUNNING)
                        emit_row_status(row_data, STATUS_RUNNING)
                        row_can_retry_after_failure = (
                            row_attempt < row_retry_count
                            and is_retry_enabled(auto_retry_enabled)
                            and not should_stop_before_next_row()
                        )

                        row_result = process_row(
                            page,
                            row_data,
                            download_dir,
                            log_dir,
                            diagnostic_events,
                            timeout_ms,
                            retry_count,
                            status_callback=lambda message, item_session_id=session_id, item_row_label=row_label, item_prefix=prefix: emit_row_progress(
                                item_session_id,
                                item_row_label,
                                item_prefix,
                                message,
                            ),
                            force_stop_requested=force_stop_requested,
                            filename_pattern=filename_pattern,
                            direct_report_count=direct_report_count,
                            task_status_callback=task_status_callback,
                            defer_country_failures_for_retry=row_can_retry_after_failure,
                        )
                        saved_files_text = row_result.saved_files_text()
                        with file_lock:
                            log_success_row(row_data, saved_files_text, log_dir)
                            update_processing_status(input_excel_path, log_dir, row_data, STATUS_SUCCESS, saved_file=saved_files_text)
                        emit_row_status(row_data, STATUS_SUCCESS, row_result.saved_files)
                        mark_result(current_index, True)
                        clear_wait_status(session_id)
                        emit_status(f"{prefix} 성공: {row_result.saved_file_names_text()}")

                    except AutomationAborted as exc:
                        clear_wait_status(session_id)
                        error_message = str(exc) or "사용자 강제종료로 처리 중단"
                        with file_lock:
                            log_failed_row(row_data, error_message, log_dir)
                            update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                        emit_row_status(row_data, STATUS_FAILED)
                        mark_result(current_index, False)
                        emit_status(f"{prefix} 처리실패 기록 후 강제종료합니다: {error_message}")
                        break

                    except Exception as exc:
                        clear_wait_status(session_id)
                        can_retry = (
                            row_attempt < row_retry_count
                            and is_retry_enabled(auto_retry_enabled)
                            and not should_stop_before_next_row()
                        )
                        message, error_message = build_failure_error_message(
                            page,
                            row_data,
                            log_dir,
                            diagnostic_events,
                            exc,
                            suffix=retry_suffix(row_attempt, final=not can_retry),
                        )
                        if can_retry:
                            should_requeue = True
                            next_attempt = row_attempt + 1
                            with file_lock:
                                update_processing_status(input_excel_path, log_dir, row_data, STATUS_RETRY_PENDING, error_message=error_message)
                            emit_row_status(row_data, STATUS_RETRY_PENDING)
                            next_retry_item = (current_index, row_data, next_attempt)
                            mark_retry_pending_failure(current_index)
                            emit_status(f"{prefix} 실패: {message}")
                            emit_status(f"{prefix} 바로 자동 재시도합니다 ({next_attempt}/{row_retry_count}).")
                            emit_progress(f"{prefix} 자동 재시도 중")
                        else:
                            with file_lock:
                                log_failed_row(row_data, error_message, log_dir)
                                update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                            emit_row_status(row_data, STATUS_FAILED)
                            mark_result(current_index, False)
                            emit_status(f"{prefix} 최종 실패: {message}")

                    finally:
                        if from_queue:
                            row_queue.task_done()
                        emit_progress(f"{prefix} 다음 행 준비 중")
                        if should_requeue or next_retry_item is not None or (not should_stop_before_next_row() and not row_queue.empty()):
                            try:
                                if page_is_closed(page) or browser_is_closed(browser):
                                    recover_worker_page("브라우저 또는 페이지가 닫혔습니다.")
                                else:
                                    reset_for_next_row(page)
                            except Exception as exc:
                                emit_status(f"[세션 {session_id}] 다음 행 준비 중 페이지 초기화 실패, 새로 접속합니다: {exc}")
                                recover_worker_page(str(exc))

            finally:
                try:
                    finish_worker(context, session_id)
                finally:
                    close_browser_context(context, browser)

    worker_errors: list[str] = []
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = [executor.submit(run_worker, session_id) for session_id in range(1, worker_count + 1)]
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as exc:
                worker_errors.append(str(exc))
                emit_status(f"병렬 세션 오류: {exc}")

    intentionally_stopped = should_stop_before_next_row()
    with counter_lock:
        uncompleted_numbers = [
            row_number
            for row_number in range(1, total + 1)
            if row_number not in completed_row_numbers
        ]

    if uncompleted_numbers and not intentionally_stopped:
        error_message = (
            "병렬 세션이 중단되어 이 행은 처리되지 않았습니다. "
            "실패 행만 재시도로 다시 실행해주세요."
        )
        with file_lock:
            for row_number in uncompleted_numbers:
                row_data = rows[row_number - 1]
                log_failed_row(row_data, error_message, log_dir)
                update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                # 대시보드에도 실패를 반영한다(세션 없이 종료된 행이므로 session_id=0).
                emit_row_status_for(0, row_data, STATUS_FAILED)
                mark_result(row_number, False)
        emit_status(f"병렬 세션 중단으로 미처리 {len(uncompleted_numbers)}건을 실패로 기록했습니다.")

    emit_progress("완료")
    with counter_lock:
        result = {
            "total": total,
            "success": success_count,
            "failed": failed_count,
            "stopped": intentionally_stopped,
            "force_stopped": combined_force_stop_requested(),
        }

    if worker_errors and result["success"] + result["failed"] == 0:
        raise RuntimeError("모든 병렬 세션이 시작 또는 처리 중 실패했습니다: " + " / ".join(worker_errors))

    return result


def run_automation(
    input_excel_path: str | Path,
    download_dir: str | Path = DEFAULT_DOWNLOAD_DIR,
    headless: bool = DEFAULT_HEADLESS,
    *,
    log_dir: str | Path = DEFAULT_LOG_DIR,
    state_path: str | Path = DEFAULT_STATE_PATH,
    timeout_ms: int = TIMEOUT_MS,
    retry_count: int = GENERATION_RETRY_COUNT,
    use_storage_state: bool = False,
    save_storage_state: bool = False,
    retry_failed_only: bool = False,
    wait_for_manual_login: bool = False,
    status_callback: StatusCallback | None = None,
    progress_callback: ProgressCallback | None = None,
    stop_requested: StopFlag | None = None,
    force_stop_requested: StopFlag | None = None,
    parallel_sessions: int = 1,
    row_retry_count: int = DEFAULT_ROW_RETRY_COUNT,
    direct_report_count: int = DEFAULT_DIRECT_REPORT_COUNT,
    auto_retry_enabled: RetryFlag | None = None,
    filename_pattern: str = "",
    report_mode: str = REPORT_MODE_DIRECT,
    recommend_then_direct: bool = False,
    row_status_callback: RowStatusCallback | None = None,
    rows_ready_callback: Callable[[list[dict[str, Any]]], None] | None = None,
    task_status_callback: TaskStatusCallback | None = None,
) -> dict[str, Any]:
    input_excel_path = Path(input_excel_path)
    download_dir = Path(download_dir)
    log_dir = Path(log_dir)
    state_path = Path(state_path)
    filename_pattern = normalize_filename_pattern(filename_pattern)
    validate_filename_pattern(filename_pattern)
    report_mode = normalize_report_mode(report_mode)
    direct_report_count = normalize_direct_report_count(direct_report_count)

    download_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    rows = read_failed_rows(log_dir) if retry_failed_only else read_input_excel(input_excel_path)
    for index, row in enumerate(rows, start=1):
        row["ui_key"] = f"row:{index}"
        row[SOURCE_FILE_COLUMN] = str(row.get(SOURCE_FILE_COLUMN, "") or input_excel_path)
        if retry_failed_only:
            row["report_mode"] = normalize_report_mode(row.get("report_mode", report_mode))
            raw_recommend_then_direct = str(row.get("recommend_then_direct", "")).strip()
            row["recommend_then_direct"] = truthy(raw_recommend_then_direct) if raw_recommend_then_direct else bool(recommend_then_direct)
        else:
            row["report_mode"] = report_mode
            row["recommend_then_direct"] = bool(recommend_then_direct)
        uses_recommend_link = row["report_mode"] == REPORT_MODE_RECOMMEND and truthy(row.get("recommend_then_direct", False))
        row["direct_report_count"] = direct_report_count if uses_recommend_link else DEFAULT_DIRECT_REPORT_COUNT
        row["use_task_resume"] = bool(retry_failed_only)
    if rows_ready_callback:
        rows_ready_callback(rows)
    total = len(rows)
    row_retry_count = normalize_row_retry_count(row_retry_count)
    success_count = 0
    failed_count = 0
    completed_count = 0
    last_progress_index = 0
    retry_pending_failed_numbers: set[int] = set()

    def emit_status(message: str) -> None:
        if status_callback:
            status_callback(message)
        else:
            print(message)

    def emit_progress(row_index: int | None = None, status: str = "") -> None:
        if progress_callback:
            progress_callback(
                {
                    "total": total,
                    "current": row_index or 0,
                    "success": success_count,
                    "failed": failed_count,
                    "status": status,
                }
            )

    def mark_retry_pending_failure(current_index: int) -> None:
        nonlocal failed_count
        if current_index not in retry_pending_failed_numbers:
            retry_pending_failed_numbers.add(current_index)
            failed_count += 1

    def resolve_retry_pending_failure(current_index: int, success: bool) -> bool:
        nonlocal failed_count
        if current_index not in retry_pending_failed_numbers:
            return False
        retry_pending_failed_numbers.remove(current_index)
        if success:
            failed_count = max(0, failed_count - 1)
        return True

    def emit_row_status(row_data: dict[str, Any], status: str, saved_files: list[Path] | None = None) -> None:
        if row_status_callback:
            row_index = int(row_data.get("row_index", 0))
            row_status_callback({
                "ui_key": str(row_data.get("ui_key", "") or f"row:{row_index}"),
                "row_index": row_index,
                "status": status,
                "session_id": 1,
                "product_name": str(row_data.get("product_name", "")),
                "hs_code": str(row_data.get("hs_code", "")),
                "target_country": str(row_data.get("target_country", "")),
                "report_mode": str(row_data.get("report_mode", "")),
                "recommend_then_direct": bool(row_data.get("recommend_then_direct", False)),
                "saved_files": [str(path) for path in (saved_files or [])],
                "ts": time.time(),
            })

    if retry_failed_only and total == 0:
        emit_status("재시도할 실패 행이 없습니다.")
        emit_progress(0, "완료")
        return {"total": 0, "success": 0, "failed": 0, "stopped": False, "force_stopped": False}

    if not retry_failed_only:
        initialize_processing_status(input_excel_path, log_dir, rows)

    if total == 0:
        emit_status("처리할 입력 행이 없습니다.")
        emit_progress(0, "완료")
        return {"total": 0, "success": 0, "failed": 0, "stopped": False, "force_stopped": False}

    parallel_sessions = normalize_parallel_sessions(parallel_sessions)
    if parallel_sessions > 1:
        return run_parallel_automation(
            input_excel_path,
            download_dir,
            headless,
            log_dir=log_dir,
            state_path=state_path,
            rows=rows,
            timeout_ms=timeout_ms,
            retry_count=retry_count,
            use_storage_state=use_storage_state,
            save_storage_state=save_storage_state,
            retry_failed_only=retry_failed_only,
            wait_for_manual_login=wait_for_manual_login,
            parallel_sessions=parallel_sessions,
            row_retry_count=row_retry_count,
            direct_report_count=direct_report_count,
            auto_retry_enabled=auto_retry_enabled,
            filename_pattern=filename_pattern,
            status_callback=status_callback,
            progress_callback=progress_callback,
            stop_requested=stop_requested,
            force_stop_requested=force_stop_requested,
            row_status_callback=row_status_callback,
            task_status_callback=task_status_callback,
        )

    with sync_playwright() as playwright:
        try:
            browser = launch_browser(playwright, headless, emit_status)
        except RuntimeError as exc:
            error_path = write_startup_error(log_dir, str(exc))
            emit_status(f"브라우저 시작 오류를 기록했습니다: {error_path}")
            raise

        context, page, diagnostic_events = create_context_page(
            browser,
            use_storage_state=use_storage_state,
            state_path=state_path,
        )

        def recover_page(reason: str) -> None:
            nonlocal browser, context, page, diagnostic_events
            emit_status(f"브라우저 세션을 복구합니다: {reason}")
            close_browser_context(context, None)
            if browser_is_closed(browser):
                close_browser_context(None, browser)
                browser = launch_browser(playwright, headless, emit_status)
            context, page, diagnostic_events = create_context_page(
                browser,
                use_storage_state=use_storage_state,
                state_path=state_path,
            )
            open_site(page)
            ensure_initial_form(page, emit_status)

        try:
            emit_status("KOTRA 보고서 생성 페이지에 접속합니다.")
            open_site(page)
            ensure_initial_form(page, emit_status)

            if wait_for_manual_login and not headless:
                emit_status("로그인이 필요하면 브라우저에서 직접 로그인한 뒤 Enter를 누르세요.")
                input("로그인 완료 후 Enter를 누르세요: ")
                ensure_initial_form(page, emit_status)
                if save_storage_state:
                    context.storage_state(path=str(state_path))
                    emit_status(f"로그인 세션을 저장했습니다: {state_path}")

            row_queue: queue.Queue[tuple[int, dict[str, Any], int]] = queue.Queue()
            for row_number, row_data in enumerate(rows, start=1):
                row_queue.put((row_number, row_data, 0))

            next_retry_item: tuple[int, dict[str, Any], int] | None = None

            while True:
                check_force_stop(force_stop_requested)
                if stop_requested and stop_requested():
                    emit_status("중지 요청이 있어 다음 행으로 넘어가지 않습니다.")
                    break

                from_queue = False
                if next_retry_item is not None:
                    current_index, row_data, row_attempt = next_retry_item
                    next_retry_item = None
                else:
                    try:
                        current_index, row_data, row_attempt = row_queue.get_nowait()
                        from_queue = True
                    except queue.Empty:
                        break

                row_index = int(row_data["row_index"])

                row_label = f"{current_index}/{total}"
                if retry_failed_only:
                    row_label = f"{row_label}, 원본 행 {row_index}"
                if row_attempt:
                    row_label = f"{row_label}, 자동 재시도 {row_attempt}/{row_retry_count}"

                should_requeue = False
                emit_progress(completed_count, "입력 및 보고서 생성 중")
                emit_status(f"[{row_label}] 행 처리를 시작합니다.")
                clear_diagnostics(diagnostic_events)
                row_data["use_task_resume"] = bool(retry_failed_only or row_attempt > 0)
                update_processing_status(input_excel_path, log_dir, row_data, STATUS_RUNNING)
                emit_row_status(row_data, STATUS_RUNNING)

                try:
                    row_can_retry_after_failure = (
                        row_attempt < row_retry_count
                        and is_retry_enabled(auto_retry_enabled)
                        and not (stop_requested and stop_requested())
                        and not (force_stop_requested and force_stop_requested())
                    )
                    row_result = process_row(
                        page,
                        row_data,
                        download_dir,
                        log_dir,
                        diagnostic_events,
                        timeout_ms,
                        retry_count,
                        status_callback=lambda message: emit_progress(completed_count, message),
                        force_stop_requested=force_stop_requested,
                        filename_pattern=filename_pattern,
                        direct_report_count=direct_report_count,
                        task_status_callback=task_status_callback,
                        defer_country_failures_for_retry=row_can_retry_after_failure,
                    )
                    saved_files_text = row_result.saved_files_text()
                    log_success_row(row_data, saved_files_text, log_dir)
                    resolve_retry_pending_failure(current_index, success=True)
                    success_count += 1
                    completed_count += 1
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_SUCCESS, saved_file=saved_files_text)
                    emit_row_status(row_data, STATUS_SUCCESS, row_result.saved_files)
                    emit_status(f"[{row_label}] 성공: {row_result.saved_file_names_text()}")

                except AutomationAborted as exc:
                    error_message = str(exc) or "사용자 강제종료로 처리 중단"
                    was_retry_pending_failure = resolve_retry_pending_failure(current_index, success=False)
                    if not was_retry_pending_failure:
                        failed_count += 1
                    completed_count += 1
                    log_failed_row(row_data, error_message, log_dir)
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                    emit_row_status(row_data, STATUS_FAILED)
                    emit_status(f"[{row_label}] 처리실패 기록 후 강제종료합니다: {error_message}")
                    break

                except Exception as exc:
                    can_retry = (
                        row_attempt < row_retry_count
                        and is_retry_enabled(auto_retry_enabled)
                        and not (stop_requested and stop_requested())
                        and not (force_stop_requested and force_stop_requested())
                    )
                    message, error_message = build_failure_error_message(
                        page,
                        row_data,
                        log_dir,
                        diagnostic_events,
                        exc,
                        suffix=retry_suffix(row_attempt, final=not can_retry),
                    )

                    if can_retry:
                        should_requeue = True
                        next_attempt = row_attempt + 1
                        update_processing_status(input_excel_path, log_dir, row_data, STATUS_RETRY_PENDING, error_message=error_message)
                        emit_row_status(row_data, STATUS_RETRY_PENDING)
                        next_retry_item = (current_index, row_data, next_attempt)
                        mark_retry_pending_failure(current_index)
                        emit_status(f"[{row_label}] 실패: {message}")
                        emit_status(f"[{row_label}] 바로 자동 재시도합니다 ({next_attempt}/{row_retry_count}).")
                        emit_progress(completed_count, "자동 재시도 중")
                    else:
                        was_retry_pending_failure = resolve_retry_pending_failure(current_index, success=False)
                        if not was_retry_pending_failure:
                            failed_count += 1
                        completed_count += 1
                        log_failed_row(row_data, error_message, log_dir)
                        update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                        emit_row_status(row_data, STATUS_FAILED)
                        emit_status(f"[{row_label}] 최종 실패: {message}")

                finally:
                    if from_queue:
                        row_queue.task_done()

                last_progress_index = completed_count
                emit_progress(completed_count, "다음 행 준비 중")
                if (
                    (should_requeue or next_retry_item is not None or not row_queue.empty())
                    and not (stop_requested and stop_requested())
                    and not (force_stop_requested and force_stop_requested())
                ):
                    try:
                        if page_is_closed(page) or browser_is_closed(browser):
                            recover_page("브라우저 또는 페이지가 닫혔습니다.")
                        else:
                            reset_for_next_row(page)
                    except Exception as exc:
                        emit_status(f"다음 행 준비 중 페이지 초기화 실패, 새로 접속합니다: {exc}")
                        recover_page(str(exc))

            emit_progress(last_progress_index, "완료")

            if save_storage_state:
                try:
                    context.storage_state(path=str(state_path))
                except Exception as exc:
                    emit_status(f"브라우저 세션 저장을 건너뜁니다: {exc}")

        finally:
            close_browser_context(context, browser)

    stopped = bool(stop_requested and stop_requested()) or bool(force_stop_requested and force_stop_requested())
    return {
        "total": total,
        "success": success_count,
        "failed": failed_count,
        "stopped": stopped,
        "force_stopped": bool(force_stop_requested and force_stop_requested()),
    }


def row_identity(row_data: dict[str, Any]) -> tuple[str, ...]:
    return (
        str(row_data.get(SOURCE_FILE_COLUMN, "")).strip(),
        normalize_report_mode(row_data.get("report_mode", "")),
        str(truthy(row_data.get("recommend_then_direct", False))),
        str(row_data.get("row_index", "")).strip(),
        str(row_data.get("hs_code", "")).strip(),
        str(row_data.get("product_name", "")).strip(),
        normalize_country_key(row_data.get("target_country", "")),
        normalize_country_key(row_data.get("excluded_countries", "")),
    )


def log_record_identity(record: dict[str, Any], fallback_key: str = "") -> tuple[str, ...]:
    row_index = str(record.get("row_index", "")).strip() or fallback_key
    return (
        str(record.get(SOURCE_FILE_COLUMN, "")).strip(),
        normalize_report_mode(record.get("report_mode", "")),
        str(truthy(record.get("recommend_then_direct", False))),
        row_index,
        str(record.get("hs_code", "")).strip(),
        str(record.get("product_name", "")).strip(),
        normalize_country_key(record.get("target_country", "")),
        normalize_country_key(record.get("excluded_countries", "")),
    )


def processing_status_path(log_dir: str | Path) -> Path:
    return Path(log_dir) / PROCESSING_STATUS_FILENAME


def report_tasks_path(log_dir: str | Path) -> Path:
    return Path(log_dir) / REPORT_TASKS_FILENAME


def initialize_processing_status(input_excel_path: str | Path, log_dir: str | Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return

    status_path = processing_status_path(log_dir)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status_rows = []

    for row_data in rows:
        status_row = build_processing_status_row(input_excel_path, row_data)
        status_row[STATUS_COLUMN] = STATUS_PENDING
        status_row[STATUS_AT_COLUMN] = now
        status_row[SAVED_FILE_COLUMN] = ""
        status_row[ERROR_COLUMN] = ""
        status_rows.append(status_row)
        row_data["process_status"] = STATUS_PENDING
        row_data["saved_file"] = ""
        row_data["error_message"] = ""

    write_processing_status_rows(status_path, status_rows)


def update_processing_status(
    input_excel_path: str | Path,
    log_dir: str | Path,
    row_data: dict[str, Any],
    status: str,
    *,
    saved_file: str | Path = "",
    error_message: str = "",
) -> None:
    status_path = processing_status_path(log_dir)
    existing_rows = read_processing_status_rows(status_path)
    existing_by_key = {row_identity(row): row for row in existing_rows}
    ordered_keys = [row_identity(row) for row in existing_rows]
    next_status_row = build_processing_status_row(input_excel_path, row_data)
    key = row_identity(next_status_row)
    status_row = existing_by_key.get(key, next_status_row)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    status_row.update(next_status_row)
    status_row[STATUS_COLUMN] = status
    status_row[STATUS_AT_COLUMN] = now
    status_row[SAVED_FILE_COLUMN] = str(saved_file) if saved_file else str(row_data.get("saved_file", ""))
    status_row[ERROR_COLUMN] = error_message

    if key not in existing_by_key:
        ordered_keys.append(key)
    existing_by_key[key] = status_row
    write_processing_status_rows(status_path, [existing_by_key[item_key] for item_key in ordered_keys if item_key in existing_by_key])

    row_data["process_status"] = status
    row_data["saved_file"] = str(saved_file) if saved_file else ""
    row_data["error_message"] = error_message


def build_processing_status_row(input_excel_path: str | Path, row_data: dict[str, Any]) -> dict[str, str]:
    return {
        "report_mode": str(row_data.get("report_mode", "")),
        SOURCE_FILE_COLUMN: str(row_data.get(SOURCE_FILE_COLUMN, "") or Path(input_excel_path)),
        "recommend_then_direct": str(row_data.get("recommend_then_direct", "")),
        "direct_report_count": str(row_data.get("direct_report_count", "")),
        "row_index": str(row_data.get("row_index", "")),
        "hs_code": str(row_data.get("hs_code", "")),
        "product_name": str(row_data.get("product_name", "")),
        "export_scale": str(row_data.get("export_scale", "")),
        "export_experience": str(row_data.get("export_experience", "")),
        "target_country": str(row_data.get("target_country", "")),
        "excluded_countries": str(row_data.get("excluded_countries", "")),
        "recommended_countries": str(row_data.get("recommended_countries", "")),
        "final_target_countries": str(row_data.get("final_target_countries", "")),
        "recommendation_report_file": str(row_data.get("recommendation_report_file", "")),
        "direct_report_files": str(row_data.get("direct_report_files", "")),
        STATUS_COLUMN: str(row_data.get("process_status", "")),
        STATUS_AT_COLUMN: "",
        SAVED_FILE_COLUMN: str(row_data.get("saved_file", "")),
        ERROR_COLUMN: str(row_data.get("error_message", "")),
    }


def read_processing_status_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []

    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    rows: list[dict[str, str]] = []
    for record in df.to_dict(orient="records"):
        rows.append({column: str(record.get(column, "")) for column in PROCESSING_STATUS_COLUMNS})
    return rows


def write_processing_status_rows(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(
        [{column: row.get(column, "") for column in PROCESSING_STATUS_COLUMNS} for row in rows],
        columns=PROCESSING_STATUS_COLUMNS,
    )
    df.to_excel(path, index=False)


def update_report_task_status(
    log_dir: str | Path,
    row_data: dict[str, Any],
    task_type: str,
    country: str = "",
    status: str = "",
    *,
    saved_file: str | Path = "",
    error_message: str = "",
) -> None:
    with REPORT_TASKS_LOCK:
        path = report_tasks_path(log_dir)
        existing_rows = read_report_task_rows(path)
        existing_by_key = {report_task_identity(row): row for row in existing_rows}
        ordered_keys = [report_task_identity(row) for row in existing_rows]
        task_row = build_report_task_row(row_data, task_type, country)
        key = report_task_identity(task_row)
        current_row = existing_by_key.get(key, task_row)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        current_row.update(task_row)
        current_row["status"] = status
        if saved_file:
            current_row["saved_file"] = str(saved_file)
        elif status == STATUS_SUCCESS:
            current_row["saved_file"] = str(current_row.get("saved_file", ""))
        else:
            current_row["saved_file"] = ""
        current_row["recommended_countries"] = str(row_data.get("recommended_countries", current_row.get("recommended_countries", "")))
        current_row["final_target_countries"] = str(row_data.get("final_target_countries", current_row.get("final_target_countries", "")))
        current_row["error_message"] = error_message
        current_row["updated_at"] = now

        if key not in existing_by_key:
            ordered_keys.append(key)
        existing_by_key[key] = current_row
        write_report_task_rows(path, [existing_by_key[item_key] for item_key in ordered_keys if item_key in existing_by_key])


def completed_report_task(
    log_dir: str | Path,
    row_data: dict[str, Any],
    task_type: str,
    country: str = "",
) -> dict[str, str] | None:
    if not truthy(row_data.get("use_task_resume", False)):
        return None

    path = report_tasks_path(log_dir)
    task_row = build_report_task_row(row_data, task_type, country)
    key = report_task_identity(task_row)
    for row in reversed(read_report_task_rows(path)):
        if report_task_identity(row) != key:
            continue
        if str(row.get("status", "")) != STATUS_SUCCESS:
            return None
        saved_file = str(row.get("saved_file", "")).strip()
        if not saved_file or not Path(saved_file).exists():
            return None
        return row
    return None


def build_report_task_row(row_data: dict[str, Any], task_type: str, country: str = "") -> dict[str, str]:
    return {
        SOURCE_FILE_COLUMN: str(row_data.get(SOURCE_FILE_COLUMN, "")),
        "report_mode": str(row_data.get("report_mode", "")),
        "recommend_then_direct": str(row_data.get("recommend_then_direct", "")),
        "direct_report_count": str(row_data.get("direct_report_count", "")),
        "row_index": str(row_data.get("row_index", "")),
        "hs_code": str(row_data.get("hs_code", "")),
        "product_name": str(row_data.get("product_name", "")),
        "target_country": str(row_data.get("target_country", "")),
        "excluded_countries": str(row_data.get("excluded_countries", "")),
        "task_type": str(task_type),
        "country": normalize_single_country(country),
        "status": "",
        "saved_file": "",
        "recommended_countries": str(row_data.get("recommended_countries", "")),
        "final_target_countries": str(row_data.get("final_target_countries", "")),
        "error_message": "",
        "updated_at": "",
    }


def report_task_identity(row: dict[str, Any]) -> tuple[str, ...]:
    return (
        str(row.get(SOURCE_FILE_COLUMN, "")).strip(),
        normalize_report_mode(row.get("report_mode", "")),
        str(truthy(row.get("recommend_then_direct", False))),
        str(row.get("row_index", "")).strip(),
        str(row.get("hs_code", "")).strip(),
        str(row.get("product_name", "")).strip(),
        normalize_country_key(row.get("target_country", "")),
        normalize_country_key(row.get("excluded_countries", "")),
        str(row.get("task_type", "")).strip(),
        normalize_country_key(row.get("country", "")),
    )


def read_report_task_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []

    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    rows: list[dict[str, str]] = []
    for record in df.to_dict(orient="records"):
        rows.append({column: str(record.get(column, "")) for column in REPORT_TASK_COLUMNS})
    return rows


def write_report_task_rows(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(
        [{column: row.get(column, "") for column in REPORT_TASK_COLUMNS} for row in rows],
        columns=REPORT_TASK_COLUMNS,
    )
    df.to_excel(path, index=False)


def join_path_values(paths: list[Path]) -> str:
    return "; ".join(str(path) for path in paths)


def read_input_excel(input_excel_path: str | Path) -> list[dict[str, Any]]:
    path = Path(input_excel_path)
    if not path.exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {path}")

    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    rows: list[dict[str, Any]] = []

    for idx, record in enumerate(df.to_dict(orient="records"), start=1):
        data_values = [value for column, value in record.items() if str(column).strip() not in STATUS_COLUMNS]
        if not any(str(value).strip() for value in data_values):
            continue

        row = {}
        for key in FIELD_MAPPING.keys():
            raw_value = get_source_value(record, key)
            if key == "hs_code" and not raw_value:
                raw_value = get_source_value(record, "hs_code_10")
            if not raw_value and key in {"export_scale", "export_experience"}:
                raw_value = find_category_value(record)
            row[key] = normalize_field_value(key, raw_value)

        if not is_processable_input_row(row):
            continue

        row_index_value = normalize_row_index(get_source_value(record, "row_index"), idx)
        row["row_index"] = row_index_value
        row["company_name"] = normalize_field_value("company_name", get_source_value(record, "company_name"))
        row["business_number"] = normalize_field_value("business_number", get_source_value(record, "business_number"))
        row["excel_row_number"] = idx + 1
        row["process_status"] = ""
        row["saved_file"] = ""
        row["error_message"] = ""
        row["recommended_countries"] = ""
        row["final_target_countries"] = ""
        row["recommendation_report_file"] = ""
        row["direct_report_files"] = ""
        rows.append(row)

    return rows


def is_processable_input_row(row: dict[str, Any]) -> bool:
    return any(str(row.get(key, "")).strip() for key in ("hs_code", "product_name", "export_scale", "export_experience"))


def normalize_row_index(value: Any, fallback_index: int) -> int:
    text = str(value or "").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    if re.fullmatch(r"\d+", text):
        return int(text)
    return fallback_index


def read_failed_rows(log_dir: str | Path) -> list[dict[str, Any]]:
    log_dir = Path(log_dir)
    path = log_dir / "failed_rows.xlsx"
    if not path.exists():
        return []

    latest_success_at = read_latest_success_times(log_dir)
    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    records = df.to_dict(orient="records")
    rows: list[dict[str, Any]] = []
    seen_rows: set[tuple[str, ...]] = set()

    for fallback_index, record in reversed(list(enumerate(records, start=1))):
        row_index_text = str(record.get("row_index", "")).strip()
        dedupe_key = log_record_identity(record, f"fallback-{fallback_index}")
        if dedupe_key in seen_rows:
            continue

        failed_at = str(record.get("failed_at", "")).strip()
        success_at = latest_success_at.get(dedupe_key, "")
        if success_at and (not failed_at or success_at >= failed_at):
            seen_rows.add(dedupe_key)
            continue

        seen_rows.add(dedupe_key)

        row = {}
        for key in FIELD_MAPPING.keys():
            raw_value = get_source_value(record, key)
            row[key] = normalize_field_value(key, raw_value)

        row["report_mode"] = normalize_report_mode(str(record.get("report_mode", "")).strip())
        row["recommend_then_direct"] = str(record.get("recommend_then_direct", "")).strip()
        row[SOURCE_FILE_COLUMN] = str(record.get(SOURCE_FILE_COLUMN, "")).strip()
        row["row_index"] = int(row_index_text) if row_index_text.isdigit() else fallback_index
        row["company_name"] = normalize_field_value("company_name", get_source_value(record, "company_name"))
        row["business_number"] = normalize_field_value("business_number", get_source_value(record, "business_number"))
        row["recommended_countries"] = str(record.get("recommended_countries", "")).strip()
        row["final_target_countries"] = str(record.get("final_target_countries", "")).strip()
        row["recommendation_report_file"] = str(record.get("recommendation_report_file", "")).strip()
        row["direct_report_files"] = str(record.get("direct_report_files", "")).strip()
        rows.append(row)

    return list(reversed(rows))


def read_latest_success_times(log_dir: Path) -> dict[tuple[str, ...], str]:
    path = log_dir / "success_log.xlsx"
    if not path.exists():
        return {}

    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    latest: dict[tuple[str, str, str, str], str] = {}
    for fallback_index, record in enumerate(df.to_dict(orient="records"), start=1):
        key = log_record_identity(record, f"fallback-{fallback_index}")
        completed_at = str(record.get("completed_at", "")).strip()
        if not key[0] or not completed_at:
            continue
        if completed_at > latest.get(key, ""):
            latest[key] = completed_at
    return latest


def get_source_value(record: dict[str, Any], target_key: str) -> str:
    if target_key in record:
        value = str(record.get(target_key, "")).strip()
        if value:
            return value

    normalized_record = {normalize_column_name(column): value for column, value in record.items()}
    aliases = SOURCE_COLUMN_ALIASES.get(target_key, [target_key])
    empty_value = ""

    for alias in aliases:
        alias_key = normalize_column_name(alias)
        if alias_key in normalized_record:
            value = str(normalized_record[alias_key]).strip()
            if value:
                return value
            empty_value = value

    for alias in aliases:
        alias_key = normalize_column_name(alias)
        for column_key, value in normalized_record.items():
            if alias_key and alias_key in column_key:
                text = str(value).strip()
                if text:
                    return text
                empty_value = text

    return empty_value


def normalize_column_name(value: Any) -> str:
    return re.sub(r"[\s\-_()/:\n]+", "", str(value or "")).lower()


def find_category_value(record: dict[str, Any]) -> str:
    for value in record.values():
        text = str(value or "").strip()
        if text in EXPORT_SCALE_CATEGORY_MAP or text in EXPORT_EXPERIENCE_CATEGORY_MAP:
            return text
    return ""


def normalize_field_value(field_name: str, value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)

    if field_name == "hs_code":
        return normalize_hs_code(text)

    if field_name == "product_name":
        return text.replace("\\n", " ").strip()

    if field_name == "export_scale":
        return normalize_export_scale(text)

    if field_name == "export_experience":
        normalized = text.upper()
        if normalized in {"O", "X"}:
            return normalized
        mapped_value = EXPORT_EXPERIENCE_CATEGORY_MAP.get(text)
        if mapped_value is not None:
            return mapped_value
        compact_map = {str(key).replace(" ", ""): mapped for key, mapped in EXPORT_EXPERIENCE_CATEGORY_MAP.items()}
        return compact_map.get(text.replace(" ", ""), text)

    if field_name in {"target_country", "excluded_countries"}:
        return normalize_target_country(text)

    return text


def normalize_export_scale(value: str) -> str:
    mapped_value = EXPORT_SCALE_CATEGORY_MAP.get(value)
    if mapped_value:
        return mapped_value

    keyword_scale = export_scale_from_category_keyword(value)
    if keyword_scale:
        return keyword_scale

    amount_scale = export_scale_from_numeric_amount(value)
    if amount_scale:
        return amount_scale

    return value


def export_scale_from_category_keyword(value: str) -> str | None:
    compact = re.sub(r"\s+", "", str(value or ""))
    if not compact:
        return None

    keyword_candidates = [
        ("선도기업", "선도"),
        ("선도", "선도"),
        ("성장기업", "성장"),
        ("성장", "성장"),
        ("유망기업", "유망"),
        ("유망", "유망"),
        ("초보기업", "초보"),
        ("초보", "초보"),
        ("내수기업", "내수"),
        ("내수", "내수"),
        ("수출액없음", "내수"),
    ]
    for keyword, category in keyword_candidates:
        if keyword in compact:
            return EXPORT_SCALE_CATEGORY_MAP[category]
    return None


def export_scale_from_numeric_amount(value: str) -> str | None:
    amount = parse_numeric_amount(value)
    if amount is None:
        return None

    if amount <= 0:
        return "내수기업 (수출액 없음)"
    if amount < Decimal("100000"):
        return "초보기업 ($1 ~ $99,999)"
    if amount < Decimal("1000000"):
        return "유망기업 ($100,000 ~ $999,999)"
    if amount < Decimal("10000000"):
        return "성장기업 ($1,000,000 ~ $9,999,999)"
    return "선도기업 ($10,000,000 ~)"


def parse_numeric_amount(value: str) -> Decimal | None:
    text = str(value or "").strip()
    if not text:
        return None

    cleaned = text
    for token in ("USD", "usd", "US$", "us$", "달러", "불"):
        cleaned = cleaned.replace(token, "")
    cleaned = re.sub(r"[,\s$]", "", cleaned)

    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", cleaned):
        return None

    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def normalize_hs_code(value: str) -> str:
    value = value.strip()
    if re.fullmatch(r"\d+\.0", value):
        value = value.split(".", 1)[0]

    digits = re.sub(r"\D", "", value)
    if len(digits) >= 6:
        return digits[:6]
    return digits if digits else value


def validate_hs_code(value: Any) -> None:
    text = str(value or "").strip()
    if not re.fullmatch(r"\d{6}", text):
        raise ValueError(f"HS CODE는 6자리 숫자여야 합니다. 현재 값: {text}")


def normalize_target_country(value: str) -> str:
    # 분리/정규화 규칙을 split_country_values 와 단일화한다.
    # '-' 같은 대시 플레이스홀더('희망 국가 없음')와 중복 국가는 여기서 제거된다.
    return join_country_values(split_country_values(value))


def safe_filename(text: Any) -> str:
    value = str(text).strip()
    value = re.sub(r"[\\/:*?\"<>|]", "_", value)
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"_+", "_", value)
    return value[:80] or "empty"
