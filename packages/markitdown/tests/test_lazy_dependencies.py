"""The import boundary for optional converter dependencies."""

import subprocess
import sys
from pathlib import Path

import pytest

TEST_FILES = Path(__file__).parent / "test_files"


def _run(code: str) -> str:
    return subprocess.run(
        [sys.executable, "-c", code],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def test_importing_markitdown_does_not_load_pandas() -> None:
    output = _run("import sys, markitdown; print('pandas' in sys.modules)")
    assert output == "False"


def test_first_spreadsheet_conversion_loads_pandas() -> None:
    pytest.importorskip("pandas")
    path = TEST_FILES / "test.xlsx"
    output = _run(
        "import sys;"
        "from markitdown import MarkItDown;"
        f"MarkItDown().convert({str(path)!r});"
        "print('pandas' in sys.modules)"
    )
    assert output == "True"
