from __future__ import annotations

import re
from pathlib import Path


def test_render_gunicorn_timeout_allows_large_workbook_processing():
    render_yaml = Path("render.yaml").read_text(encoding="utf-8")
    timeout_match = re.search(r"--timeout\s+(\d+)", render_yaml)

    assert timeout_match is not None
    assert int(timeout_match.group(1)) >= 600
    assert "--workers 1" in render_yaml
