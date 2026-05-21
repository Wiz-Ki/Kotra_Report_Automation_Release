from __future__ import annotations

import argparse
from pathlib import Path

from config import (
    BASE_DIR,
    DEFAULT_CONFIG,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_INPUT_EXCEL,
    DEFAULT_LOG_DIR,
    MAX_PARALLEL_SESSIONS,
)


def parallel_session_count(value: str) -> int:
    try:
        count = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("병렬 세션 수는 숫자여야 합니다.") from exc

    if count < 1 or count > MAX_PARALLEL_SESSIONS:
        raise argparse.ArgumentTypeError(f"병렬 세션 수는 1~{MAX_PARALLEL_SESSIONS} 사이여야 합니다.")
    return count


def resolve_default_input() -> Path:
    input_xlsx = BASE_DIR / "input.xlsx"
    if input_xlsx.exists():
        return input_xlsx
    return DEFAULT_INPUT_EXCEL


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="KOTRA 수출 보고서 자동 생성 및 다운로드 프로그램")
    parser.add_argument("--gui", action="store_true", help="CustomTkinter GUI V2로 실행합니다.")
    parser.add_argument("--gui-legacy", action="store_true", help="기존 tkinter GUI로 실행합니다.")
    parser.add_argument("--input", default=str(resolve_default_input()), help="입력 엑셀 파일 경로")
    parser.add_argument("--download-dir", default=str(DEFAULT_DOWNLOAD_DIR), help="다운로드 저장 폴더")
    parser.add_argument("--log-dir", default=str(DEFAULT_LOG_DIR), help="로그 저장 폴더")
    parser.add_argument("--headless", action="store_true", help="브라우저를 백그라운드로 실행합니다.")
    parser.add_argument("--login-wait", action="store_true", help="브라우저에서 수동 작업 후 Enter를 누르고 시작합니다.")
    parser.add_argument("--no-login-wait", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--use-storage-state", action="store_true", help="state.json 브라우저 세션을 사용하고 실행 후 다시 저장합니다.")
    parser.add_argument("--no-storage-state", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--retry-failed", action="store_true", help="logs/failed_rows.xlsx에 기록된 실패 행만 다시 실행합니다.")
    parser.add_argument("--parallel-sessions", type=parallel_session_count, default=1, help=f"동시에 실행할 브라우저 세션 수(1~{MAX_PARALLEL_SESSIONS})")
    parser.add_argument("--create-template", action="store_true", help="input_template.xlsx를 생성하고 종료합니다.")
    return parser


def main() -> None:
    args = build_parser().parse_args()

    if args.create_template:
        from template import create_input_template

        template_path = create_input_template()
        print(f"입력 템플릿을 생성했습니다: {template_path}")
        return

    if args.gui:
        from gui_v2 import run_gui

        run_gui()
        return

    if args.gui_legacy:
        from gui import run_gui

        run_gui()
        return

    input_path = Path(args.input)
    if not args.retry_failed:
        from template import create_input_template

        create_input_template(DEFAULT_INPUT_EXCEL)

        if not input_path.exists():
            print(f"입력 파일이 없어 기본 템플릿을 사용합니다: {DEFAULT_INPUT_EXCEL}")
            input_path = DEFAULT_INPUT_EXCEL

    from automation import run_automation

    use_storage_state = args.use_storage_state and not args.no_storage_state

    result = run_automation(
        input_excel_path=input_path,
        download_dir=args.download_dir,
        headless=args.headless or bool(DEFAULT_CONFIG["headless"]),
        log_dir=args.log_dir,
        use_storage_state=use_storage_state,
        save_storage_state=use_storage_state,
        retry_failed_only=args.retry_failed,
        wait_for_manual_login=args.login_wait,
        parallel_sessions=args.parallel_sessions,
    )

    print("작업이 완료되었습니다.")
    print(f"전체: {result['total']}건 / 성공: {result['success']}건 / 실패: {result['failed']}건")


if __name__ == "__main__":
    main()
