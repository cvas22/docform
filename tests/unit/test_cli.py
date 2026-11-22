import json
import os
import subprocess
import sys
from pathlib import Path

from docx import Document

from conftest import body_text

ROOT = Path(__file__).resolve().parents[2]


def run_cli(*args):
    return subprocess.run(
        [sys.executable, "-m", "docform", *map(str, args)],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
    )


def test_cli_renders_json(make_docx, tmp_path):
    tpl = make_docx(["Hello {{ name }}"])
    data = tmp_path / "data.json"
    data.write_text(json.dumps({"name": "CLI"}), encoding="utf-8")
    out = tmp_path / "out.docx"

    result = run_cli(tpl, data, out, "-q")
    assert result.returncode == 0, result.stderr
    assert body_text(Document(str(out))) == ["Hello CLI"]


def test_cli_overwrite_flag(make_docx, tmp_path):
    tpl = make_docx(["v{{ n }}"])
    data = tmp_path / "data.json"
    out = tmp_path / "out.docx"
    for n in (1, 2):
        data.write_text(json.dumps({"n": n}), encoding="utf-8")
        result = run_cli(tpl, data, out, "-o", "-q")
        assert result.returncode == 0, result.stderr
    assert body_text(Document(str(out))) == ["v2"]


def test_cli_sandbox_and_unsafe_flag(make_docx, tmp_path):
    tpl = make_docx(["{{ cycler.__init__.__globals__.os.name }}"])
    data = tmp_path / "data.json"
    data.write_text("{}", encoding="utf-8")
    out = tmp_path / "out.docx"

    result = run_cli(tpl, data, out, "-q")
    assert result.returncode == 1
    assert not out.exists()

    result = run_cli(tpl, data, out, "-q", "--unsafe")
    assert result.returncode == 0, result.stderr
    assert body_text(Document(str(out))) == [os.name]


def test_cli_missing_template_fails(tmp_path):
    data = tmp_path / "data.json"
    data.write_text("{}", encoding="utf-8")
    result = run_cli(tmp_path / "nope.docx", data, tmp_path / "out.docx", "-q")
    assert result.returncode == 1
    assert "Error" in result.stderr
    assert not (tmp_path / "out.docx").exists()
