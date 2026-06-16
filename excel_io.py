from __future__ import annotations

import errno
import time
from pathlib import Path

import pandas as pd


EXCEL_WRITE_RETRY_SECONDS = 120
EXCEL_WRITE_RETRY_INTERVAL_SECONDS = 5


class ExcelFileLockedError(PermissionError):
    pass


def write_excel_with_retry(
    df: pd.DataFrame,
    path: str | Path,
    *,
    index: bool = False,
    retry_seconds: int = EXCEL_WRITE_RETRY_SECONDS,
) -> None:
    output_path = Path(path)
    deadline = time.monotonic() + retry_seconds

    while True:
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_excel(output_path, index=index)
            return
        except (PermissionError, OSError) as exc:
            if not is_permission_denied(exc):
                raise
            if time.monotonic() >= deadline:
                raise ExcelFileLockedError(
                    f"프로그램 종료 전 로그 파일을 먼저 닫아주세요.\n{output_path}"
                ) from exc
            time.sleep(EXCEL_WRITE_RETRY_INTERVAL_SECONDS)


def is_permission_denied(exc: PermissionError | OSError) -> bool:
    return isinstance(exc, PermissionError) or getattr(exc, "errno", None) == errno.EACCES
