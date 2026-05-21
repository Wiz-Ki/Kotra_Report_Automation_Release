from __future__ import annotations

import importlib.util
import sys
import sysconfig
from pathlib import Path


def ensure_stdlib_selectors() -> None:
    """
    과거 selectors.py 파일이 있을 때 표준라이브러리 selectors와 충돌하지 않게 합니다.
    pandas, subprocess, Playwright가 표준 selectors 모듈을 안정적으로 참조해야 합니다.
    """
    existing = sys.modules.get("selectors")
    if existing and hasattr(existing, "SelectSelector"):
        return

    stdlib_path = Path(sysconfig.get_paths()["stdlib"]) / "selectors.py"
    spec = importlib.util.spec_from_file_location("selectors", stdlib_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("표준라이브러리 selectors 모듈을 찾을 수 없습니다.")

    module = importlib.util.module_from_spec(spec)
    sys.modules["selectors"] = module
    spec.loader.exec_module(module)
