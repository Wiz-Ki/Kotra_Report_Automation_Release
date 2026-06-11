from __future__ import annotations

import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import automation
import pandas as pd

from automation import (
    STATUS_RETRY_PENDING,
    STATUS_SUCCESS,
    TASK_TYPE_DIRECT,
    TASK_TYPE_RECOMMEND,
    completed_report_task,
    normalize_export_scale,
    normalize_direct_report_count,
    normalize_hs_code,
    read_report_task_rows,
    report_tasks_path,
    read_failed_rows,
    read_input_excel,
    resolve_download_save_path,
    render_filename_pattern,
    split_country_values,
    update_report_task_status,
    validate_hs_code,
)
from logger import log_failed_row


class FilenamePatternTest(unittest.TestCase):
    def test_renders_custom_filename(self) -> None:
        filename = render_filename_pattern(
            "{row_index}_{company_name}_{business_number}_{hs_code}_{product_name}_{target_country}_{datetime}",
            {
                "row_index": 3,
                "company_name": "테스트기업 A",
                "business_number": "0000000001",
                "hs_code": "330499",
                "product_name": "스킨케어",
                "target_country": "베트남",
            },
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "3_테스트기업 A_0000000001_330499_스킨케어_베트남_20260609_143012.pdf")

    def test_sanitizes_forbidden_filename_characters(self) -> None:
        filename = render_filename_pattern(
            "{target_country}_{product_name}",
            {"target_country": "미국/일본", "product_name": "화장품:기초"},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "미국_일본_화장품_기초.pdf")

    def test_sanitizes_default_site_filename_for_windows(self) -> None:
        class FakeDownload:
            suggested_filename = "미국/일본:보고서.pdf"

        with TemporaryDirectory() as tmp_dir:
            path = resolve_download_save_path(FakeDownload(), Path(tmp_dir))

        self.assertEqual(path.name, "미국_일본_보고서.pdf")

    def test_avoids_windows_reserved_filename(self) -> None:
        filename = render_filename_pattern(
            "{product_name}",
            {"product_name": "CON"},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "_CON.pdf")

    def test_truncates_long_filename_for_windows(self) -> None:
        filename = render_filename_pattern(
            "{product_name}",
            {"product_name": "가" * 300},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertLessEqual(len(filename), 244)
        self.assertTrue(filename.endswith(".pdf"))

    def test_keeps_single_pdf_suffix(self) -> None:
        filename = render_filename_pattern(
            "{site_filename}.pdf",
            {},
            suggested_filename="베트남_스킨케어(330499)_수출시장분석보고서.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "베트남_스킨케어(330499)_수출시장분석보고서.pdf")

    def test_renders_split_datetime_tokens(self) -> None:
        filename = render_filename_pattern(
            "{year}-{month}-{day}_{hour}.{minute}.{second}",
            {},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "2026-06-09_14.30.12.pdf")

    def test_rejects_unknown_token(self) -> None:
        with self.assertRaises(ValueError):
            render_filename_pattern(
                "{unknown}_{hs_code}",
                {"hs_code": "330499"},
                suggested_filename="report.pdf",
                now=datetime(2026, 6, 9, 14, 30, 12),
            )

    def test_rejects_unclosed_token(self) -> None:
        with self.assertRaises(ValueError):
            render_filename_pattern(
                "{hs_code",
                {"hs_code": "330499"},
                suggested_filename="report.pdf",
                now=datetime(2026, 6, 9, 14, 30, 12),
            )

    def test_renders_recommendation_tokens(self) -> None:
        filename = render_filename_pattern(
            "{report_mode}_{excluded_countries}_{hs_code}",
            {"report_mode": "recommend", "excluded_countries": "중국, 미국", "hs_code": "330499"},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "유망 시장 추천 보고서 생성_중국, 미국_330499.pdf")

    def test_reads_company_tokens_from_input_excel(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            pd.DataFrame(
                [
                    {
                        "연번(선택)": "7",
                        "회사명(선택)": "테스트기업 A",
                        "사업자번호(선택)": "0000000001",
                        "수출액 규모(필수: 내수/초보/유망/성장/선도 중 1개)": "내수",
                        "해당 품목 수출 경험(필수: O/X)": "X",
                        "HSCODE 6단위": "330499",
                        "수출품명(선택: 구체적으로 작성 권장)": "스킨케어",
                        "희망진출국가(직접분석: 기본 2개, 추천연동: 설정값만큼 입력 가능)": "베트남",
                    }
                ]
            ).to_excel(input_path, index=False)

            [row] = read_input_excel(input_path)

        self.assertEqual(row["company_name"], "테스트기업 A")
        self.assertEqual(row["business_number"], "0000000001")

    def test_blank_row_index_does_not_use_business_number(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            pd.DataFrame(
                [
                    {
                        "연번(선택)": "1",
                        "사업자등록번호": "1111111111",
                        "수출액 규모(필수: 내수/초보/유망/성장/선도 중 1개)": "내수",
                        "해당 품목 수출 경험(필수: O/X)": "X",
                        "HSCODE 6단위": "330499",
                        "수출품명(선택: 구체적으로 작성 권장)": "스킨케어",
                        "희망진출국가(직접분석: 기본 2개, 추천연동: 설정값만큼 입력 가능)": "베트남",
                    },
                    {
                        "연번(선택)": "",
                        "사업자등록번호": "123125989",
                        "수출액 규모(필수: 내수/초보/유망/성장/선도 중 1개)": "초보",
                        "해당 품목 수출 경험(필수: O/X)": "O",
                        "HSCODE 6단위": "220600",
                        "수출품명(선택: 구체적으로 작성 권장)": "전통 쌀 발효주(막걸리)",
                        "희망진출국가(직접분석: 기본 2개, 추천연동: 설정값만큼 입력 가능)": "중국",
                    },
                ]
            ).to_excel(input_path, index=False)

            rows = read_input_excel(input_path)

        self.assertEqual(rows[0]["row_index"], 1)
        self.assertEqual(rows[1]["row_index"], 2)
        self.assertEqual(rows[1]["business_number"], "123125989")

    def test_splits_country_values(self) -> None:
        self.assertEqual(split_country_values("베트남, 미국/일본\n베트남"), ["베트남", "미국", "일본"])

    def test_normalizes_export_scale_with_mixed_amount_text(self) -> None:
        self.assertEqual(
            normalize_export_scale("성장기업($1,000,000 이상)"),
            "성장기업 ($1,000,000 ~ $9,999,999)",
        )
        self.assertEqual(
            normalize_export_scale("전년도 기준 선도기업 금액"),
            "선도기업 ($10,000,000 ~)",
        )
        self.assertEqual(
            normalize_export_scale("내수기업(수출액 없음)"),
            "내수기업 (수출액 없음)",
        )

    def test_normalizes_ten_digit_hs_code_to_six_digits(self) -> None:
        self.assertEqual(normalize_hs_code("3304990000"), "330499")
        self.assertEqual(normalize_hs_code("330499.0"), "330499")

    def test_rejects_short_hs_code_instead_of_padding(self) -> None:
        self.assertEqual(normalize_hs_code("123"), "123")
        with self.assertRaises(ValueError):
            validate_hs_code("123")

    def test_normalizes_direct_report_count_with_limits(self) -> None:
        self.assertEqual(normalize_direct_report_count("3"), 3)
        self.assertEqual(normalize_direct_report_count("0"), 1)
        self.assertEqual(normalize_direct_report_count("999"), 5)
        self.assertEqual(normalize_direct_report_count("bad"), 2)

    def test_reads_external_request_template_shape(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "request_template.xlsx"
            pd.DataFrame(
                [
                    {
                        "연번": "1",
                        "HSCODE 10단위\n(모르는 경우 생략 가능)": "3307904000",
                        "HSCODE 6단위": "330499",
                        "수출품명": "마스크팩",
                        "수출액 규모": "성장기업 ($1,000,000~$9,999,999)",
                        "해당 품목 수출경험": "수출경험 있음",
                        "희망 진출국 (최대 2개국)\n(없는 경우 생략 가능)": "베트남, 미국",
                        "분석 제외 국가\n(없는 경우 생략 가능)": "이스라엘",
                    },
                    {"희망 진출국 (최대 2개국)\n(없는 경우 생략 가능)": "베트남"},
                    {"희망 진출국 (최대 2개국)\n(없는 경우 생략 가능)": "-"},
                ]
            ).to_excel(path, index=False)

            rows = read_input_excel(path)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["hs_code"], "330499")
        self.assertEqual(rows[0]["export_experience"], "O")
        self.assertEqual(rows[0]["target_country"], "베트남, 미국")
        self.assertEqual(rows[0]["excluded_countries"], "이스라엘")

    def test_completed_report_task_requires_resume_flag_and_existing_file(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            log_dir = Path(tmp_dir) / "logs"
            saved_file = Path(tmp_dir) / "베트남.pdf"
            saved_file.write_text("pdf", encoding="utf-8")
            row = {
                "report_mode": "direct",
                "row_index": 1,
                "hs_code": "330499",
                "product_name": "마스크팩",
                "target_country": "베트남, 미국",
            }

            update_report_task_status(log_dir, row, TASK_TYPE_DIRECT, "베트남", STATUS_SUCCESS, saved_file=saved_file)

            self.assertIsNone(completed_report_task(log_dir, row, TASK_TYPE_DIRECT, "베트남"))
            row["use_task_resume"] = True
            self.assertIsNotNone(completed_report_task(log_dir, row, TASK_TYPE_DIRECT, "베트남"))

            update_report_task_status(log_dir, row, TASK_TYPE_DIRECT, "베트남", "처리 중")
            task_rows = read_report_task_rows(report_tasks_path(log_dir))
            self.assertEqual(task_rows[-1]["saved_file"], "")

    def test_report_task_identity_keeps_excluded_countries_separate(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            log_dir = Path(tmp_dir) / "logs"
            saved_file = Path(tmp_dir) / "추천_중국제외.pdf"
            saved_file.write_text("pdf", encoding="utf-8")
            row = {
                "report_mode": "recommend",
                "recommend_then_direct": False,
                "row_index": 1,
                "hs_code": "330499",
                "product_name": "마스크팩",
                "target_country": "",
                "excluded_countries": "중국",
                "use_task_resume": True,
            }

            update_report_task_status(log_dir, row, TASK_TYPE_RECOMMEND, "", STATUS_SUCCESS, saved_file=saved_file)

            other_excluded_row = {**row, "excluded_countries": "일본"}
            self.assertIsNone(completed_report_task(log_dir, other_excluded_row, TASK_TYPE_RECOMMEND))
            self.assertIsNotNone(completed_report_task(log_dir, row, TASK_TYPE_RECOMMEND))

    def test_recommend_link_continues_remaining_direct_country_after_failure(self) -> None:
        class FakePage:
            def is_closed(self) -> bool:
                return False

        with TemporaryDirectory() as tmp_dir:
            log_dir = Path(tmp_dir) / "logs"
            save_dir = Path(tmp_dir) / "downloads"
            save_dir.mkdir()
            recommendation_file = Path(tmp_dir) / "추천.pdf"
            recommendation_file.write_text("pdf", encoding="utf-8")
            row = {
                "report_mode": "recommend",
                "recommend_then_direct": True,
                "direct_report_count": 2,
                "row_index": 2,
                "hs_code": "220600",
                "product_name": "전통 쌀 발효주(막걸리)",
                "target_country": "중국, 미국",
                "excluded_countries": "",
                "recommended_countries": "",
                "final_target_countries": "중국, 미국",
                "use_task_resume": True,
            }
            update_report_task_status(
                log_dir,
                row,
                TASK_TYPE_RECOMMEND,
                "",
                STATUS_SUCCESS,
                saved_file=recommendation_file,
            )

            def fake_process_direct_country_report(_page, _row, country, *_args, **_kwargs):
                if country == "중국":
                    raise RuntimeError("서버 오류")
                path = save_dir / f"{country}.pdf"
                path.write_text("pdf", encoding="utf-8")
                return path

            with patch.object(automation, "reset_for_next_row"), patch.object(
                automation,
                "process_direct_country_report",
                side_effect=fake_process_direct_country_report,
            ):
                with self.assertRaisesRegex(RuntimeError, "중국"):
                    automation.process_recommendation_row(
                        FakePage(),
                        row,
                        save_dir,
                        log_dir,
                        [],
                        timeout_ms=1,
                        retry_count=0,
                        status_callback=None,
                        force_stop_requested=None,
                        filename_pattern="",
                        direct_report_count=2,
                        defer_country_failures_for_retry=True,
                    )

            task_rows = read_report_task_rows(report_tasks_path(log_dir))
            statuses = {
                (item["task_type"], item["country"]): item["status"]
                for item in task_rows
            }

        self.assertEqual(statuses[(TASK_TYPE_DIRECT, "중국")], STATUS_RETRY_PENDING)
        self.assertEqual(statuses[(TASK_TYPE_DIRECT, "미국")], STATUS_SUCCESS)

    def test_failed_rows_restore_report_mode_options(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            log_dir = Path(tmp_dir) / "logs"
            log_failed_row(
                {
                    "report_mode": "recommend",
                    "recommend_then_direct": True,
                    "row_index": 1,
                    "hs_code": "330499",
                    "product_name": "마스크팩",
                    "export_scale": "성장기업 ($1,000,000 ~ $9,999,999)",
                    "export_experience": "O",
                    "target_country": "",
                    "excluded_countries": "이스라엘",
                },
                "테스트 실패",
                log_dir,
            )

            rows = read_failed_rows(log_dir)

        self.assertEqual(rows[0]["report_mode"], "recommend")
        self.assertEqual(rows[0]["recommend_then_direct"], "True")


if __name__ == "__main__":
    unittest.main()
