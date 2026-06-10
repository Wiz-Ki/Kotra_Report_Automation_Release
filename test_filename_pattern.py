from __future__ import annotations

import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from automation import normalize_export_scale, normalize_hs_code, read_input_excel, render_filename_pattern, split_country_values


class FilenamePatternTest(unittest.TestCase):
    def test_renders_custom_filename(self) -> None:
        filename = render_filename_pattern(
            "{row_index}_{hs_code}_{product_name}_{target_country}_{datetime}",
            {
                "row_index": 3,
                "hs_code": "330499",
                "product_name": "스킨케어",
                "target_country": "베트남",
            },
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "3_330499_스킨케어_베트남_20260609_143012.pdf")

    def test_sanitizes_forbidden_filename_characters(self) -> None:
        filename = render_filename_pattern(
            "{target_country}_{product_name}",
            {"target_country": "미국/일본", "product_name": "화장품:기초"},
            suggested_filename="report.pdf",
            now=datetime(2026, 6, 9, 14, 30, 12),
        )

        self.assertEqual(filename, "미국_일본_화장품_기초.pdf")

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


if __name__ == "__main__":
    unittest.main()
