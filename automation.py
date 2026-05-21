from __future__ import annotations

import re
import importlib.util
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
    DEFAULT_HEADLESS,
    DEFAULT_LOG_DIR,
    DEFAULT_STATE_PATH,
    ELEMENT_TIMEOUT_MS,
    GENERATION_RETRY_COUNT,
    KOTRA_REPORT_URL,
    PAGE_LOAD_TIMEOUT_MS,
    TIMEOUT_MS,
)
from field_mapping import (
    EXPORT_EXPERIENCE_CATEGORY_MAP,
    EXPORT_SCALE_CATEGORY_MAP,
    FIELD_MAPPING,
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
StopFlag = Callable[[], bool]

PROCESSING_STATUS_FILENAME = "processing_status.xlsx"
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
    "row_index",
    "hs_code",
    "product_name",
    "export_scale",
    "export_experience",
    "target_country",
    *STATUS_COLUMNS,
]
STATUS_PENDING = "처리 안됨"
STATUS_RUNNING = "처리 중"
STATUS_SUCCESS = "처리완료"
STATUS_FAILED = "처리실패"


class GenerationError(RuntimeError):
    def __init__(self, message: str, artifacts: str = "") -> None:
        super().__init__(message)
        self.artifacts = artifacts


class AutomationAborted(RuntimeError):
    pass


def launch_edge_browser(playwright: Any, headless: bool):
    try:
        return playwright.chromium.launch(channel="msedge", headless=headless)
    except PlaywrightError as exc:
        raise RuntimeError(
            "Microsoft Edge를 Playwright로 실행하지 못했습니다.\n"
            "- Windows VM에 Microsoft Edge가 설치되어 있는지 확인해주세요.\n"
            "- 사내 Edge 정책이 자동화 실행 또는 새 브라우저 프로필 생성을 차단할 수 있습니다.\n"
            "- Edge가 실행되지만 다운로드가 실패하면 보안/DLP 정책의 다운로드 차단 여부를 확인해주세요.\n"
            f"원본 오류: {exc}"
        ) from exc


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


def fill_form(page: Page, row_data: dict[str, Any]) -> None:
    """
    FIELD_MAPPING을 기준으로 엑셀 한 행의 데이터를 사이트 입력창에 채운다.
    """
    for column_name, field_info in FIELD_MAPPING.items():
        value = str(row_data.get(column_name, "")).strip()

        if field_info.get("required") and not value:
            raise ValueError(f"필수 입력값 누락: {column_name}")

        if not value:
            continue

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


def click_generate_button(page: Page) -> None:
    button = page.locator(SELECTORS["generate_button"]).first
    button.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    button.scroll_into_view_if_needed()
    button.wait_for(state="attached", timeout=ELEMENT_TIMEOUT_MS)
    wait_until_enabled(page, button, ELEMENT_TIMEOUT_MS)
    button.click()


def wait_until_enabled(page: Page, locator, timeout_ms: int) -> None:
    handle = locator.element_handle(timeout=timeout_ms)
    if handle is None:
        raise RuntimeError("버튼 요소를 찾지 못했습니다.")
    page.wait_for_function(
        "(el) => !el.disabled && el.getAttribute('aria-disabled') !== 'true'",
        arg=handle,
        timeout=timeout_ms,
    )


def download_button_locators(page: Page):
    locators = [page.locator(SELECTORS["download_button"]).first]
    fallback_selector = SELECTORS.get("download_button_fallback")
    if fallback_selector:
        locators.append(page.locator(fallback_selector).first)
    return locators


def wait_for_download_button(page: Page, timeout_ms: int = TIMEOUT_MS):
    deadline = datetime.now().timestamp() + timeout_ms / 1000
    candidates = download_button_locators(page)

    while datetime.now().timestamp() < deadline:
        for candidate in candidates:
            if is_visible(candidate):
                return candidate
        page.wait_for_timeout(500)

    raise PlaywrightTimeoutError(f"다운로드 버튼이 {timeout_ms // 1000}초 안에 나타나지 않았습니다.")


def wait_for_download_or_generation_error(
    page: Page,
    timeout_ms: int = TIMEOUT_MS,
    status_callback: StatusCallback | None = None,
    force_stop_requested: StopFlag | None = None,
):
    deadline = datetime.now().timestamp() + timeout_ms / 1000
    started_at = datetime.now().timestamp()
    last_status_at = 0.0
    download_buttons = download_button_locators(page)
    retry_button = page.locator(SELECTORS["retry_button"]).first
    error_text = page.locator(SELECTORS["streaming_error_text"]).first
    generate_button = page.locator(SELECTORS["generate_button"]).first

    while datetime.now().timestamp() < deadline:
        check_force_stop(force_stop_requested)
        now = datetime.now().timestamp()
        if status_callback and now - last_status_at >= 10:
            elapsed = int(now - started_at)
            status_callback(f"보고서 생성 중입니다. 경과 {elapsed}초")
            last_status_at = now

        for download_button in download_buttons:
            if is_visible(download_button):
                if status_callback:
                    status_callback("PDF 저장 버튼이 나타났습니다. 다운로드를 시작합니다.")
                return "download", download_button
        if is_visible(retry_button) or is_visible(error_text):
            if status_callback:
                status_callback("KOTRA 서버 오류 화면이 감지되었습니다.")
            return "error", retry_button
        if now - started_at > 5 and is_visible(generate_button):
            if status_callback:
                status_callback("초기 입력 화면으로 돌아온 상태가 감지되었습니다.")
            return "returned_to_form", generate_button
        page.wait_for_timeout(1000)

    raise PlaywrightTimeoutError(f"다운로드 버튼이 {timeout_ms // 1000}초 안에 나타나지 않았습니다.")


def is_visible(locator) -> bool:
    try:
        return locator.is_visible()
    except Exception:
        return False


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
        page.screenshot(path=str(screenshot_path), full_page=True)
    except Exception as exc:
        events.append(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] screenshot failed: {exc}")

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
        f"url: {page.url}",
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


def download_report(
    page: Page,
    save_path: str | Path,
    timeout_ms: int = TIMEOUT_MS,
    status_callback: StatusCallback | None = None,
    force_stop_requested: StopFlag | None = None,
) -> Path:
    save_path = Path(save_path)
    save_path.parent.mkdir(parents=True, exist_ok=True)

    status, download_button = wait_for_download_or_generation_error(page, timeout_ms, status_callback, force_stop_requested)
    if status == "error":
        raise RuntimeError("KOTRA 보고서 생성 중 서버 스트리밍 오류가 발생했습니다.")
    if status == "returned_to_form":
        raise RuntimeError("보고서 생성 중 초기 입력 화면으로 돌아왔습니다. 서버 오류 또는 사용자의 되돌아가기 동작으로 판단됩니다.")

    download_button.scroll_into_view_if_needed()
    wait_until_enabled(page, download_button, ELEMENT_TIMEOUT_MS)
    check_force_stop(force_stop_requested)

    with page.expect_download(timeout=timeout_ms) as download_info:
        download_button.click()

    download = download_info.value
    download.save_as(str(save_path))
    if status_callback:
        status_callback("PDF 다운로드가 완료되었습니다.")
    return save_path


def reset_for_next_row(page: Page) -> None:
    new_analysis_button = page.locator(SELECTORS["new_analysis_button"]).first
    if is_visible(new_analysis_button):
        new_analysis_button.scroll_into_view_if_needed()
        wait_until_enabled(page, new_analysis_button, ELEMENT_TIMEOUT_MS)
        new_analysis_button.click()
        page.locator(SELECTORS["hs_code_input"]).first.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        return

    try:
        reset_button = page.locator(SELECTORS["reset_button"]).first
        reset_button.wait_for(state="visible", timeout=5_000)
        reset_button.scroll_into_view_if_needed()
        reset_button.click()
        page.locator(SELECTORS["hs_code_input"]).first.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        open_site(page)
        page.locator(SELECTORS["hs_code_input"]).first.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)


def ensure_initial_form(page: Page, status_callback: StatusCallback | None = None) -> None:
    """
    저장된 브라우저 세션이 결과 화면을 복원해도 다음 자동화가 입력 화면에서 시작되도록 보장한다.
    """
    hs_code_input = page.locator(SELECTORS["hs_code_input"]).first
    new_analysis_button = page.locator(SELECTORS["new_analysis_button"]).first
    reset_button = page.locator(SELECTORS["reset_button"]).first

    deadline = datetime.now().timestamp() + 10
    while datetime.now().timestamp() < deadline:
        if is_visible(hs_code_input):
            return
        if is_visible(new_analysis_button):
            if status_callback:
                status_callback("이전 결과 화면이 감지되어 새 분석 화면으로 돌아갑니다.")
            new_analysis_button.scroll_into_view_if_needed()
            wait_until_enabled(page, new_analysis_button, ELEMENT_TIMEOUT_MS)
            new_analysis_button.click()
            hs_code_input.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            return
        if is_visible(reset_button):
            if status_callback:
                status_callback("입력 화면을 초기화합니다.")
            reset_button.scroll_into_view_if_needed()
            reset_button.click()
            hs_code_input.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            return
        page.wait_for_timeout(300)

    if status_callback:
        status_callback("초기 입력 화면이 보이지 않아 KOTRA 페이지를 새로 엽니다.")
    open_site(page)
    hs_code_input.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)


def retry_generation(page: Page) -> None:
    retry_button = page.locator(SELECTORS["retry_button"]).first
    retry_button.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    retry_button.scroll_into_view_if_needed()
    wait_until_enabled(page, retry_button, ELEMENT_TIMEOUT_MS)
    retry_button.click()


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
) -> Path:
    check_force_stop(force_stop_requested)
    select_direct_country_analysis(page)
    check_force_stop(force_stop_requested)
    fill_form(page, row_data)
    check_force_stop(force_stop_requested)
    click_generate_button(page)

    last_error: Exception | None = None
    for attempt in range(retry_count + 1):
        try:
            if status_callback:
                status_callback("보고서 생성 요청을 보냈습니다.")
            return download_report(page, save_path, timeout_ms, status_callback, force_stop_requested)
        except RuntimeError as exc:
            last_error = exc
            if "초기 입력 화면" in str(exc):
                artifacts = save_failure_artifacts(
                    page,
                    row_data,
                    log_dir,
                    str(exc),
                    diagnostic_events,
                    suffix="returned_to_form",
                )
                raise GenerationError(str(exc), artifacts) from exc

            if "서버 스트리밍 오류" not in str(exc) or attempt >= retry_count:
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
) -> dict[str, int]:
    input_excel_path = Path(input_excel_path)
    download_dir = Path(download_dir)
    log_dir = Path(log_dir)
    state_path = Path(state_path)

    download_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    rows = read_failed_rows(log_dir) if retry_failed_only else read_input_excel(input_excel_path)
    total = len(rows)
    success_count = 0
    failed_count = 0
    last_progress_index = 0

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

    if retry_failed_only and total == 0:
        emit_status("재시도할 실패 행이 없습니다.")
        emit_progress(0, "완료")
        return {"total": 0, "success": 0, "failed": 0}

    if not retry_failed_only:
        initialize_processing_status(input_excel_path, log_dir, rows)

    if total == 0:
        emit_status("처리할 입력 행이 없습니다.")
        emit_progress(0, "완료")
        return {"total": 0, "success": 0, "failed": 0}

    with sync_playwright() as playwright:
        emit_status("Microsoft Edge를 실행합니다.")
        try:
            browser = launch_edge_browser(playwright, headless)
        except RuntimeError as exc:
            error_path = write_startup_error(log_dir, str(exc))
            emit_status(f"브라우저 시작 오류를 기록했습니다: {error_path}")
            raise

        context_kwargs: dict[str, Any] = {"accept_downloads": True}

        if use_storage_state and state_path.exists():
            context_kwargs["storage_state"] = str(state_path)

        context = browser.new_context(**context_kwargs)
        page = context.new_page()
        page.set_default_navigation_timeout(PAGE_LOAD_TIMEOUT_MS)
        page.set_default_timeout(ELEMENT_TIMEOUT_MS)
        diagnostic_events = setup_page_diagnostics(page)

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

            for current_index, row_data in enumerate(rows, start=1):
                check_force_stop(force_stop_requested)
                row_index = int(row_data["row_index"])
                if stop_requested and stop_requested():
                    emit_status("중지 요청이 있어 다음 행으로 넘어가지 않습니다.")
                    break

                emit_progress(current_index, "입력 및 보고서 생성 중")
                row_label = f"{current_index}/{total}"
                if retry_failed_only:
                    row_label = f"{row_label}, 원본 행 {row_index}"
                emit_status(f"[{row_label}] 행 처리를 시작합니다.")
                clear_diagnostics(diagnostic_events)
                update_processing_status(input_excel_path, log_dir, row_data, STATUS_RUNNING)

                save_path = build_download_path(download_dir, row_data)

                try:
                    saved_file = process_row(
                        page,
                        row_data,
                        save_path,
                        log_dir,
                        diagnostic_events,
                        timeout_ms,
                        retry_count,
                        status_callback=lambda message, current=current_index: emit_progress(current, message),
                        force_stop_requested=force_stop_requested,
                    )
                    log_success_row(row_data, saved_file, log_dir)
                    success_count += 1
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_SUCCESS, saved_file=saved_file)
                    emit_status(f"[{row_label}] 성공: {saved_file.name}")

                except PlaywrightTimeoutError as exc:
                    failed_count += 1
                    message = f"제한 시간 안에 필요한 요소를 찾지 못했습니다: {exc}"
                    artifacts = save_failure_artifacts(page, row_data, log_dir, message, diagnostic_events)
                    error_message = f"{message} / diagnostics: {artifacts}"
                    log_failed_row(row_data, error_message, log_dir)
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                    emit_status(f"[{row_label}] 실패: {message}")

                except AutomationAborted:
                    emit_status("강제종료 요청으로 작업을 즉시 중단합니다.")
                    break

                except GenerationError as exc:
                    failed_count += 1
                    error_message = f"{exc} / diagnostics: {exc.artifacts}"
                    log_failed_row(row_data, error_message, log_dir)
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                    emit_status(f"[{row_label}] 실패: {exc}")

                except Exception as exc:
                    failed_count += 1
                    artifacts = save_failure_artifacts(page, row_data, log_dir, str(exc), diagnostic_events)
                    error_message = f"{exc} / diagnostics: {artifacts}"
                    log_failed_row(row_data, error_message, log_dir)
                    update_processing_status(input_excel_path, log_dir, row_data, STATUS_FAILED, error_message=error_message)
                    emit_status(f"[{row_label}] 실패: {exc}")

                finally:
                    last_progress_index = current_index
                    emit_progress(current_index, "다음 행 준비 중")
                    if (
                        current_index != total
                        and not (stop_requested and stop_requested())
                        and not (force_stop_requested and force_stop_requested())
                    ):
                        try:
                            reset_for_next_row(page)
                        except Exception as exc:
                            emit_status(f"다음 행 준비 중 페이지 초기화 실패, 새로 접속합니다: {exc}")
                            open_site(page)
                            ensure_initial_form(page, emit_status)

            emit_progress(last_progress_index, "완료")

            if save_storage_state:
                context.storage_state(path=str(state_path))

        finally:
            context.close()
            browser.close()

    return {"total": total, "success": success_count, "failed": failed_count}


def row_identity(row_data: dict[str, Any]) -> tuple[str, str, str, str]:
    return (
        str(row_data.get("row_index", "")).strip(),
        str(row_data.get("hs_code", "")).strip(),
        str(row_data.get("product_name", "")).strip(),
        str(row_data.get("target_country", "")).strip(),
    )


def processing_status_path(log_dir: str | Path) -> Path:
    return Path(log_dir) / PROCESSING_STATUS_FILENAME


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
    key = row_identity(row_data)
    status_row = existing_by_key.get(key, build_processing_status_row(input_excel_path, row_data))
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    status_row.update(build_processing_status_row(input_excel_path, row_data))
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
        SOURCE_FILE_COLUMN: str(Path(input_excel_path)),
        "row_index": str(row_data.get("row_index", "")),
        "hs_code": str(row_data.get("hs_code", "")),
        "product_name": str(row_data.get("product_name", "")),
        "export_scale": str(row_data.get("export_scale", "")),
        "export_experience": str(row_data.get("export_experience", "")),
        "target_country": str(row_data.get("target_country", "")),
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
            if not raw_value and key in {"export_scale", "export_experience"}:
                raw_value = find_category_value(record)
            row[key] = normalize_field_value(key, raw_value)

        row["row_index"] = idx
        row["excel_row_number"] = idx + 1
        row["process_status"] = ""
        row["saved_file"] = ""
        row["error_message"] = ""
        rows.append(row)

    return rows


def read_failed_rows(log_dir: str | Path) -> list[dict[str, Any]]:
    log_dir = Path(log_dir)
    path = log_dir / "failed_rows.xlsx"
    if not path.exists():
        return []

    latest_success_at = read_latest_success_times(log_dir)
    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    records = df.to_dict(orient="records")
    rows: list[dict[str, Any]] = []
    seen_row_indexes: set[str] = set()

    for fallback_index, record in reversed(list(enumerate(records, start=1))):
        row_index_text = str(record.get("row_index", "")).strip()
        dedupe_key = row_index_text or f"fallback-{fallback_index}"
        if dedupe_key in seen_row_indexes:
            continue

        failed_at = str(record.get("failed_at", "")).strip()
        success_at = latest_success_at.get(dedupe_key, "")
        if success_at and (not failed_at or success_at >= failed_at):
            seen_row_indexes.add(dedupe_key)
            continue

        seen_row_indexes.add(dedupe_key)

        row = {}
        for key in FIELD_MAPPING.keys():
            raw_value = get_source_value(record, key)
            row[key] = normalize_field_value(key, raw_value)

        row["row_index"] = int(row_index_text) if row_index_text.isdigit() else fallback_index
        rows.append(row)

    return list(reversed(rows))


def read_latest_success_times(log_dir: Path) -> dict[str, str]:
    path = log_dir / "success_log.xlsx"
    if not path.exists():
        return {}

    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    latest: dict[str, str] = {}
    for record in df.to_dict(orient="records"):
        row_index = str(record.get("row_index", "")).strip()
        completed_at = str(record.get("completed_at", "")).strip()
        if not row_index or not completed_at:
            continue
        if completed_at > latest.get(row_index, ""):
            latest[row_index] = completed_at
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
        return EXPORT_EXPERIENCE_CATEGORY_MAP.get(text, text)

    if field_name == "target_country":
        return normalize_target_country(text)

    return text


def normalize_export_scale(value: str) -> str:
    mapped_value = EXPORT_SCALE_CATEGORY_MAP.get(value)
    if mapped_value:
        return mapped_value

    amount_scale = export_scale_from_numeric_amount(value)
    if amount_scale:
        return amount_scale

    return value


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
    return digits.zfill(6) if digits else value


def normalize_target_country(value: str) -> str:
    value = value.replace("/", ", ")
    value = re.sub(r"\s*,\s*", ", ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" ,")


def build_download_path(download_dir: str | Path, row_data: dict[str, Any]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    row_index = int(row_data.get("row_index", 0))
    hs_code = safe_filename(row_data.get("hs_code", ""))
    product_name = safe_filename(row_data.get("product_name", ""))
    target_country = safe_filename(str(row_data.get("target_country", "")).replace(",", ""))
    filename = f"{row_index:03d}_{hs_code}_{product_name}_{target_country}_{timestamp}.pdf"
    return Path(download_dir) / filename


def safe_filename(text: Any) -> str:
    value = str(text).strip()
    value = re.sub(r"[\\/:*?\"<>|]", "_", value)
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"_+", "_", value)
    return value[:80] or "empty"
