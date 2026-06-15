from __future__ import annotations

import argparse
from pathlib import Path

from config import (
    BASE_DIR,
    DEFAULT_CONFIG,
    DEFAULT_DIRECT_REPORT_COUNT,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_INPUT_EXCEL,
    DEFAULT_LOG_DIR,
    DEFAULT_ROW_RETRY_COUNT,
    MAX_DIRECT_REPORT_COUNT,
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


def direct_report_count(value: str) -> int:
    try:
        count = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("보고서 생성 개수는 숫자여야 합니다.") from exc

    if count < 1 or count > MAX_DIRECT_REPORT_COUNT:
        raise argparse.ArgumentTypeError(f"보고서 생성 개수는 1~{MAX_DIRECT_REPORT_COUNT} 사이여야 합니다.")
    return count


def resolve_default_input() -> Path:
    return BASE_DIR / "input.xlsx"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="KOTRA 수출 보고서 자동 생성 및 다운로드 프로그램")
    parser.add_argument("--gui", action="store_true", help="CustomTkinter GUI V2로 실행합니다.")
    parser.add_argument("--input", default=str(resolve_default_input()), help="입력 엑셀 파일 경로")
    parser.add_argument("--download-dir", default=str(DEFAULT_DOWNLOAD_DIR), help="다운로드 저장 폴더")
    parser.add_argument("--log-dir", default=str(DEFAULT_LOG_DIR), help="로그 저장 폴더")
    parser.add_argument("--headless", action="store_true", help="브라우저를 백그라운드로 실행합니다.")
    parser.add_argument("--no-headless", action="store_true", help="기본 백그라운드 실행을 끄고 브라우저 화면을 표시합니다.")
    parser.add_argument("--login-wait", action="store_true", help="브라우저에서 수동 작업 후 Enter를 누르고 시작합니다.")
    parser.add_argument("--no-login-wait", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--use-storage-state", action="store_true", help="state.json 브라우저 세션을 사용하고 실행 후 다시 저장합니다.")
    parser.add_argument("--no-storage-state", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--retry-failed", action="store_true", help="logs/failed_rows.xlsx에 기록된 실패 행만 다시 실행합니다.")
    parser.add_argument("--resume", action="store_true", help="logs/processing_status.xlsx 기준으로 이미 완료된 행은 건너뛰고 다시 실행합니다.")
    parser.add_argument("--no-auto-retry", action="store_true", help="행 처리 실패 시 기본 1회 자동 재시도를 사용하지 않습니다.")
    parser.add_argument("--parallel-sessions", type=parallel_session_count, default=1, help=f"동시에 실행할 브라우저 세션 수(1~{MAX_PARALLEL_SESSIONS})")
    parser.add_argument(
        "--direct-report-count",
        type=direct_report_count,
        default=DEFAULT_DIRECT_REPORT_COUNT,
        help=f"추천 연동 시 생성할 수출시장 분석보고서 수(1~{MAX_DIRECT_REPORT_COUNT}, 기본 {DEFAULT_DIRECT_REPORT_COUNT})",
    )
    parser.add_argument(
        "--report-mode",
        choices=["direct", "recommend"],
        default="direct",
        help="생성할 보고서 방식: direct=수출시장 분석, recommend=유망 시장 추천",
    )
    parser.add_argument("--recommend-then-direct", action="store_true", help="유망 시장 추천 보고서 생성 후 추천 국가로 수출시장 분석 보고서까지 생성합니다.")
    parser.add_argument(
        "--filename-pattern",
        default="",
        help="저장 파일명 커스텀 패턴. 예: {row_index}_{hs_code}_{product_name}_{datetime}",
    )
    parser.add_argument("--create-template", action="store_true", help="input_template.xlsx를 생성하고 종료합니다.")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.retry_failed and args.resume:
        print("--retry-failed와 --resume은 함께 사용할 수 없습니다. 하나만 선택해주세요.")
        return 1

    if args.create_template:
        from template import create_input_template

        template_path = create_input_template()
        print(f"입력 템플릿을 생성했습니다: {template_path}")
        return 0

    if args.gui:
        from gui_v2 import run_gui

        run_gui()
        return 0

    input_path = Path(args.input)
    if not args.retry_failed:
        if not input_path.exists():
            print(f"입력 엑셀 파일을 찾을 수 없습니다: {input_path}")
            print("input_template.xlsx를 복사해 input.xlsx로 작성하거나, --input 옵션으로 파일을 지정해주세요.")
            print("새 템플릿이 필요하면 run_cli.bat --create-template 을 실행하세요.")
            return 1

    from automation import run_automation

    use_storage_state = args.use_storage_state and not args.no_storage_state

    result = run_automation(
        input_excel_path=input_path,
        download_dir=args.download_dir,
        headless=(False if args.no_headless else (args.headless or bool(DEFAULT_CONFIG["headless"]))),
        log_dir=args.log_dir,
        use_storage_state=use_storage_state,
        save_storage_state=use_storage_state,
        retry_failed_only=args.retry_failed,
        resume_incomplete=args.resume,
        wait_for_manual_login=args.login_wait,
        parallel_sessions=args.parallel_sessions,
        row_retry_count=0 if args.no_auto_retry else DEFAULT_ROW_RETRY_COUNT,
        direct_report_count=args.direct_report_count,
        filename_pattern=args.filename_pattern,
        report_mode=args.report_mode,
        recommend_then_direct=args.recommend_then_direct,
    )

    print("작업이 완료되었습니다.")
    skipped_text = f" / 건너뜀: {result.get('skipped', 0)}건" if result.get("skipped", 0) else ""
    print(f"전체: {result['total']}건 / 성공: {result['success']}건 / 실패: {result['failed']}건{skipped_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
