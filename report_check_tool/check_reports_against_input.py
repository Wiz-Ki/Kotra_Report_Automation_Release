from __future__ import annotations

import argparse
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd


OUTPUT_FILENAME = "report_check_result.xlsx"
REPORT_SUFFIX = "수출시장분석보고서"

SOURCE_COLUMN_ALIASES = {
    "row_index": ["row_index", "연번", "연번(선택)", "순번", "번호", "NO", "No", "no"],
    "hs_code": ["hs_code", "HS CODE", "HSCODE", "HS코드"],
    "product_name": ["product_name", "수출품명", "품목명"],
    "target_country": [
        "target_country",
        "희망진출국가",
        "희망 진출 국가",
        "진출국가",
        "국가명",
        "국가",
        "수출국",
        "수출국가",
    ],
}

RESULT_COLUMNS = [
    "row_index",
    "excel_row_number",
    "hs_code",
    "product_name",
    "target_country",
    "점검결과",
    "매칭파일",
    "비고",
]


def main() -> int:
    args = build_parser().parse_args()
    folder = Path(args.folder).resolve()
    input_path = Path(args.input).resolve() if args.input else find_input_excel(folder)
    report_dir = Path(args.report_dir).resolve() if args.report_dir else folder
    output_path = Path(args.output).resolve() if args.output else folder / OUTPUT_FILENAME

    result = check_reports_against_input(input_path, report_dir, output_path, recursive=args.recursive)
    summary = result["summary"]

    print("보고서 파일 점검이 완료되었습니다.")
    print(f"입력 엑셀: {input_path}")
    print(f"보고서 폴더: {report_dir}")
    print(
        f"전체 {summary['total']}건 / 생성 확인 {summary['matched']}건 / "
        f"누락 {summary['missing']}건 / 중복 {summary['duplicate']}건 / "
        f"입력과 매칭 안 되는 PDF {summary['extra_files']}건"
    )
    print(f"결과 파일: {output_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="입력 엑셀과 생성된 PDF 보고서 폴더를 대조합니다.")
    parser.add_argument("--folder", default=".", help="엑셀과 PDF를 같이 둔 폴더")
    parser.add_argument("--input", default="", help="입력 엑셀 파일 경로. 생략하면 폴더 안의 엑셀 1개를 자동 선택합니다.")
    parser.add_argument("--report-dir", default="", help="PDF 보고서 폴더. 생략하면 --folder를 사용합니다.")
    parser.add_argument("--output", default="", help=f"결과 엑셀 경로. 기본값: {OUTPUT_FILENAME}")
    parser.add_argument("--recursive", action="store_true", help="보고서 폴더 하위 폴더의 PDF까지 검사합니다.")
    return parser


def check_reports_against_input(
    input_path: Path,
    report_dir: Path,
    output_path: Path,
    *,
    recursive: bool = False,
) -> dict[str, Any]:
    if not input_path.exists():
        raise FileNotFoundError(f"입력 엑셀 파일을 찾을 수 없습니다: {input_path}")
    if not report_dir.exists():
        raise FileNotFoundError(f"보고서 폴더를 찾을 수 없습니다: {report_dir}")

    expected_rows = read_input_rows(input_path)
    report_files = list_report_files(report_dir, recursive=recursive)
    file_infos = [parse_report_file(path) for path in report_files]
    match_index = build_match_index(file_infos)

    used_files: set[Path] = set()
    result_rows = []
    for row in expected_rows:
        matches = find_matches(row, match_index)
        for path in matches:
            used_files.add(path)
        result_rows.append(build_result_row(row, matches))

    extra_files = [info for info in file_infos if info["path"] not in used_files]
    result_df = pd.DataFrame(result_rows, columns=RESULT_COLUMNS)
    extra_df = pd.DataFrame(
        [
            {
                "파일명": info["path"].name,
                "파일경로": str(info["path"]),
                "추출_HS_CODE": info["hs_code"],
                "추출_품목명": info["product_name"],
                "추출_국가": info["target_country"],
                "비고": info["note"],
            }
            for info in extra_files
        ]
    )
    summary = build_summary(result_df, extra_df)
    write_result(output_path, result_df, extra_df, summary)
    return {"summary": summary, "rows": result_rows, "extra_files": extra_files}


def find_input_excel(folder: Path) -> Path:
    candidates = [
        path
        for path in sorted(folder.glob("*.xlsx"))
        if not path.name.startswith("~$")
        and path.name != OUTPUT_FILENAME
        and "report_check_result" not in path.stem
    ]
    if not candidates:
        raise FileNotFoundError(f"폴더 안에서 입력 엑셀 파일을 찾지 못했습니다: {folder}")
    if len(candidates) > 1:
        names = "\n".join(f"- {path.name}" for path in candidates)
        raise RuntimeError(
            "엑셀 파일이 2개 이상입니다. 하나만 남기거나 --input으로 지정해주세요.\n"
            f"{names}"
        )
    return candidates[0]


def read_input_rows(input_path: Path) -> list[dict[str, str]]:
    df = pd.read_excel(input_path, dtype=str, keep_default_na=False)
    rows: list[dict[str, str]] = []

    for fallback_index, record in enumerate(df.to_dict(orient="records"), start=1):
        if not any(str(value).strip() for value in record.values()):
            continue

        row = {
            "row_index": normalize_row_index(get_source_value(record, "row_index"), fallback_index),
            "excel_row_number": str(fallback_index + 1),
            "hs_code": normalize_hs_code(get_source_value(record, "hs_code")),
            "product_name": normalize_text(get_source_value(record, "product_name")),
            "target_country": normalize_text(get_source_value(record, "target_country")),
        }
        rows.append(row)

    return rows


def get_source_value(record: dict[str, Any], target_key: str) -> str:
    normalized_record = {normalize_column_name(column): value for column, value in record.items()}
    aliases = SOURCE_COLUMN_ALIASES.get(target_key, [target_key])

    for alias in aliases:
        alias_key = normalize_column_name(alias)
        if alias_key in normalized_record:
            text = str(normalized_record[alias_key]).strip()
            if text:
                return text

    for alias in aliases:
        alias_key = normalize_column_name(alias)
        for column_key, value in normalized_record.items():
            if alias_key and alias_key in column_key:
                text = str(value).strip()
                if text:
                    return text

    return ""


def list_report_files(report_dir: Path, *, recursive: bool) -> list[Path]:
    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted(path for path in report_dir.glob(pattern) if path.is_file())


def parse_report_file(path: Path) -> dict[str, Any]:
    name = path.name
    current = re.match(
        rf"^(?P<country>[^_]+)_(?P<product>.+)\((?P<hs_code>[^()]+)\)_{REPORT_SUFFIX}(?:\s*\(\d+\))?\.pdf$",
        name,
        flags=re.IGNORECASE,
    )
    if current:
        return {
            "path": path,
            "hs_code": normalize_hs_code(current.group("hs_code")),
            "product_name": normalize_text(current.group("product")),
            "target_country": normalize_text(current.group("country")),
            "note": "",
        }

    legacy = re.match(
        r"^\d+_(?P<hs_code>[^_]+)_(?P<product>.+)_(?P<country>[^_]+)_\d{8}_\d{6}(?:\s*\(\d+\))?\.pdf$",
        name,
        flags=re.IGNORECASE,
    )
    if legacy:
        return {
            "path": path,
            "hs_code": normalize_hs_code(legacy.group("hs_code")),
            "product_name": normalize_text(legacy.group("product")),
            "target_country": normalize_text(legacy.group("country")),
            "note": "이전 파일명 형식",
        }

    return {
        "path": path,
        "hs_code": "",
        "product_name": "",
        "target_country": "",
        "note": "파일명에서 정보를 추출하지 못함",
    }


def build_match_index(file_infos: list[dict[str, Any]]) -> dict[str, dict[tuple[str, ...], list[Path]]]:
    index: dict[str, dict[tuple[str, ...], list[Path]]] = {
        "full": defaultdict(list),
        "hs_country": defaultdict(list),
    }
    for info in file_infos:
        hs_code = info["hs_code"]
        country_key = compare_key(info["target_country"])
        product_key = compare_key(info["product_name"])
        if hs_code and country_key and product_key:
            index["full"][(hs_code, country_key, product_key)].append(info["path"])
        if hs_code and country_key:
            index["hs_country"][(hs_code, country_key)].append(info["path"])
    return index


def find_matches(row: dict[str, str], match_index: dict[str, dict[tuple[str, ...], list[Path]]]) -> list[Path]:
    hs_code = row["hs_code"]
    country_key = compare_key(row["target_country"])
    product_key = compare_key(row["product_name"])

    full_matches = match_index["full"].get((hs_code, country_key, product_key), [])
    if full_matches:
        return sorted(full_matches)

    hs_country_matches = match_index["hs_country"].get((hs_code, country_key), [])
    if len(hs_country_matches) == 1:
        return sorted(hs_country_matches)
    return []


def build_result_row(row: dict[str, str], matches: list[Path]) -> dict[str, str]:
    if not matches:
        result = "누락"
        note = "입력 행과 매칭되는 PDF를 찾지 못했습니다."
    elif len(matches) == 1:
        result = "생성 확인"
        note = ""
    else:
        result = "중복 생성"
        note = f"{len(matches)}개 PDF가 매칭되었습니다."

    return {
        **row,
        "점검결과": result,
        "매칭파일": "\n".join(str(path) for path in matches),
        "비고": note,
    }


def build_summary(result_df: pd.DataFrame, extra_df: pd.DataFrame) -> dict[str, int]:
    counts = result_df["점검결과"].value_counts().to_dict()
    return {
        "total": int(len(result_df)),
        "matched": int(counts.get("생성 확인", 0)),
        "missing": int(counts.get("누락", 0)),
        "duplicate": int(counts.get("중복 생성", 0)),
        "extra_files": int(len(extra_df)),
    }


def write_result(output_path: Path, result_df: pd.DataFrame, extra_df: pd.DataFrame, summary: dict[str, int]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    summary_df = pd.DataFrame(
        [
            {"항목": "전체 입력 행", "건수": summary["total"]},
            {"항목": "생성 확인", "건수": summary["matched"]},
            {"항목": "누락", "건수": summary["missing"]},
            {"항목": "중복 생성", "건수": summary["duplicate"]},
            {"항목": "입력과 매칭 안 되는 PDF", "건수": summary["extra_files"]},
        ]
    )

    with pd.ExcelWriter(output_path) as writer:
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        result_df.to_excel(writer, index=False, sheet_name="all")
        result_df[result_df["점검결과"] == "누락"].to_excel(writer, index=False, sheet_name="missing")
        result_df[result_df["점검결과"] == "중복 생성"].to_excel(writer, index=False, sheet_name="duplicates")
        extra_df.to_excel(writer, index=False, sheet_name="extra_pdfs")


def normalize_column_name(value: Any) -> str:
    return re.sub(r"[\s\-_()/:\n]+", "", str(value or "")).lower()


def normalize_row_index(value: Any, fallback_index: int) -> str:
    text = str(value or "").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text if text else str(fallback_index)


def normalize_hs_code(value: Any) -> str:
    text = str(value or "").strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(6) if digits else text


def normalize_text(value: Any) -> str:
    text = str(value or "").strip()
    text = text.replace("\\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compare_key(value: Any) -> str:
    text = normalize_text(value).casefold()
    text = re.sub(r"[\s\\/:*?\"<>|_(),.·ㆍ-]+", "", text)
    return text


if __name__ == "__main__":
    raise SystemExit(main())
