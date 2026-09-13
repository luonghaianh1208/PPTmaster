"""Tách chuỗi có ~chỉ số dưới~, ^chỉ số trên^ và **in đậm** thành các đoạn để dựng Word."""

from __future__ import annotations

import re

PLAIN = ""
BOLD = "bold"
SUB = "sub"
SUP = "sup"

_TOKEN = re.compile(r"\*\*(.+?)\*\*|~([^~]+)~|\^([^\^]+)\^")


def split_runs(text: str) -> list[tuple[str, str]]:
    """Trả về danh sách (nội dung, kiểu); kiểu thuộc PLAIN, BOLD, SUB, SUP."""
    runs: list[tuple[str, str]] = []
    position = 0
    for match in _TOKEN.finditer(text):
        if match.start() > position:
            runs.append((text[position:match.start()], PLAIN))
        bold, sub, sup = match.group(1), match.group(2), match.group(3)
        if bold is not None:
            runs.append((bold, BOLD))
        elif sub is not None:
            runs.append((sub, SUB))
        else:
            runs.append((sup, SUP))
        position = match.end()
    if position < len(text):
        runs.append((text[position:], PLAIN))
    return runs or [("", PLAIN)]


def plain_text(text: str) -> str:
    """Bỏ hết dấu đánh dấu, chỉ còn chữ — dùng khi đo độ dài để xếp cột đáp án."""
    return "".join(chunk for chunk, _ in split_runs(text))
