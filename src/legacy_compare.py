from __future__ import annotations

import marshal
import sys
import types
from functools import lru_cache
from pathlib import Path


@lru_cache(maxsize=1)
def load_legacy_compare_core() -> types.ModuleType:
    if sys.version_info[:2] != (3, 14):
        raise RuntimeError("The original compare_core bytecode requires Python 3.14.")
    pyc_path = Path(__file__).resolve().parents[1] / "legacy" / "compare_core.pyc"
    module = types.ModuleType("legacy_compare_core")
    code = marshal.loads(pyc_path.read_bytes())
    exec(code, module.__dict__)
    return module


def compare_sheets(old_file_path: Path, new_file_path: Path, output_dir: Path) -> dict:
    return load_legacy_compare_core().compare_sheets(old_file_path, new_file_path, output_dir)
