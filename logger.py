from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from compat import ensure_stdlib_selectors

ensure_stdlib_selectors()

import pandas as pd


SUCCESS_COLUMNS = [
    "row_index",
    "hs_code",
    "product_name",
    "export_scale",
    "export_experience",
    "target_country",
    "saved_file",
    "completed_at",
]

FAILED_COLUMNS = [
    "row_index",
    "hs_code",
    "product_name",
    "export_scale",
    "export_experience",
    "target_country",
    "error_message",
    "failed_at",
]


def ensure_log_dir(log_dir: str | Path) -> Path:
    path = Path(log_dir)
    path.mkdir(parents=True, exist_ok=True)
    return path


def _append_row(excel_path: Path, row: dict[str, Any], columns: list[str]) -> None:
    if excel_path.exists():
        df = pd.read_excel(excel_path, dtype=str, keep_default_na=False)
    else:
        df = pd.DataFrame(columns=columns)

    df = pd.concat([df, pd.DataFrame([{column: row.get(column, "") for column in columns}])], ignore_index=True)
    df.to_excel(excel_path, index=False)


def log_success_row(row_data: dict[str, Any], saved_file: str | Path, log_dir: str | Path) -> None:
    log_dir_path = ensure_log_dir(log_dir)
    row = {
        **_base_row(row_data),
        "saved_file": str(saved_file),
        "completed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    _append_row(log_dir_path / "success_log.xlsx", row, SUCCESS_COLUMNS)


def log_failed_row(row_data: dict[str, Any], error_message: str, log_dir: str | Path) -> None:
    log_dir_path = ensure_log_dir(log_dir)
    row = {
        **_base_row(row_data),
        "error_message": error_message,
        "failed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    _append_row(log_dir_path / "failed_rows.xlsx", row, FAILED_COLUMNS)


def _base_row(row_data: dict[str, Any]) -> dict[str, Any]:
    return {
        "row_index": row_data.get("row_index", ""),
        "hs_code": row_data.get("hs_code", ""),
        "product_name": row_data.get("product_name", ""),
        "export_scale": row_data.get("export_scale", ""),
        "export_experience": row_data.get("export_experience", ""),
        "target_country": row_data.get("target_country", ""),
    }
