from __future__ import annotations

import unittest
from datetime import datetime

from automation import render_filename_pattern


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


if __name__ == "__main__":
    unittest.main()
