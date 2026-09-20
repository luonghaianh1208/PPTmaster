"""Chạy mô hình JavaScript qua Node để kiểm số. Máy không có Node thì trả về None, không báo lỗi."""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

RUNTIME_DIR = Path(__file__).resolve().parent / "runtime"
RUNNER = RUNTIME_DIR / "chay_node.js"
KHUNG_JS = RUNTIME_DIR / "khung.js"
TIMEOUT_S = 60


class CheckError(Exception):
    """Node có chạy nhưng mô hình hỏng: lỗi cú pháp, ném lỗi, hoặc quá thời gian."""


def find_node() -> str | None:
    return shutil.which("node")


def _run(model, mode: str, stdin: str = ""):
    node = find_node()
    if node is None:
        return None
    try:
        proc = subprocess.run(
            [node, str(RUNNER), str(KHUNG_JS), str(model.js_path), str(model.json_path), mode],
            input=stdin, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=TIMEOUT_S,
        )
    except subprocess.TimeoutExpired as exc:
        raise CheckError(f"Mô hình chạy quá {TIMEOUT_S} giây, có thể bị lặp vô hạn") from exc
    except OSError as exc:
        raise CheckError(f"Không chạy được Node: {exc}") from exc
    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except (IndexError, json.JSONDecodeError) as exc:
        raise CheckError(f"Node không trả về JSON: {(proc.stderr or proc.stdout).strip()[:300]}") from exc
    if isinstance(data, dict) and "loi" in data:
        raise CheckError(f"Mô hình lỗi khi chạy: {data['loi']}")
    return data


def bang_kiem(model) -> dict | None:
    """{'tong', 'dat', 'truot': [{'dong', 'ma', 'mong', 'duoc'}]} hoặc None khi máy không có Node."""
    return _run(model, "bang-kiem")


def luoi(model, diem: list[dict]) -> list[dict] | None:
    """Giá trị các đại lượng đo tại từng bộ tham số; None khi máy không có Node."""
    return _run(model, "luoi", json.dumps(diem, ensure_ascii=False))
