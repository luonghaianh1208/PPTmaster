"""Cắt phần của một bài ra khỏi file SGK đã chuyển sang Markdown.

Mục đích: không bao giờ phải nạp cả quyển SGK (19–22 MB) vào ngữ cảnh.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import difflib
import re

from .frameworks import normalise

MAX_KB = 200
_BAI_RE = re.compile(r"^\**\s*bai\s+\d+\b")
_LESSON_NUM_RE = re.compile(r"bai\s+(\d+)")
_PUNCT_RE = re.compile(r"[^0-9a-z\s]")
_CHUONG_RE = re.compile(r"^chuong(\s|\d)")


class SgkError(Exception):
    """Không cắt được phần của bài khỏi file SGK."""

    def __init__(self, message: str, fix: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.fix = fix


def headings(lines: list[str]) -> list[tuple[int, str]]:
    """Trả về (chỉ số dòng, tiêu đề) cho mọi tiêu đề Markdown và mọi dòng mở đầu bằng 'Bài <số>'."""
    found: list[tuple[int, str]] = []
    for index, raw in enumerate(lines):
        text = raw.strip()
        if not text:
            continue
        if text.startswith("#"):
            found.append((index, text.lstrip("# ").strip()))
        elif _BAI_RE.match(normalise(text)):
            found.append((index, text.strip("*# ").strip()))
    return found


def _squash(text: str) -> str:
    """Bỏ mọi dấu câu, chỉ giữ chữ/số/khoảng trắng — để 'Bài 5. Ammonia' khớp 'BÀI 5: AMMONIA'."""
    return " ".join(_PUNCT_RE.sub(" ", text).split())


def _lesson_number(title: str) -> int | None:
    match = _LESSON_NUM_RE.search(normalise(title))
    return int(match.group(1)) if match else None


def _contains(title: str, target: str) -> bool:
    """'bai 5' khớp 'bai 5. ammonia' nhưng không khớp 'bai 50. he sinh thai'."""
    start = title.find(target)
    while start != -1:
        after = title[start + len(target):start + len(target) + 1]
        if not (target[-1:].isdigit() and after.isdigit()):
            return True
        start = title.find(target, start + 1)
    return False


def _level(line: str) -> int:
    """Số dấu # đầu dòng; dòng 'Bài <số>' thường (không có #) là 0."""
    text = line.strip()
    return len(text) - len(text.lstrip("#"))


def _adjacent_relevant(lines: list[str], index: int, step: int) -> str | None:
    """Tìm dòng liền kề (theo hướng step) không rỗng và không phải dòng chương, đã chuẩn hoá."""
    i = index + step
    while 0 <= i < len(lines):
        text = lines[i].strip()
        if not text:
            i += step
            continue
        norm = normalise(text)
        if _CHUONG_RE.match(norm):
            i += step
            continue
        return norm
    return None


def _neighbor_lesson_number(lines: list[str], index: int, step: int) -> int | None:
    """Số bài của dòng liền kề (bỏ dòng trống/dòng chương) nếu đó là tiêu đề 'Bài <số>'."""
    norm = _adjacent_relevant(lines, index, step)
    if norm is None or _BAI_RE.match(norm) is None:
        return None
    match = _LESSON_NUM_RE.search(norm)
    return int(match.group(1)) if match else None


def _looks_like_toc_entry(lines: list[str], index: int, number: int | None) -> bool:
    """Dòng 'Bài <số>' thường có phải một mục mục lục không?

    Chỉ coi là mục lục khi dòng liền kề (bỏ qua dòng trống và dòng chương) là một tiêu đề
    'Bài <số>' khác mang đúng số liền trước hoặc liền sau số bài này — mục lục liệt kê các
    bài theo đúng thứ tự liên tiếp. Dòng lặp lại đầu trang bên trong nội dung bài, hay một
    dòng bài tập đánh số ('Bài 1.', 'Bài 2.'...) đứng cạnh tiêu đề thật, mang số bất kỳ
    không liền kề, nên không bị nhầm là mục lục.
    Giới hạn đã biết, chưa xử lý: sách chỉ có một bài với mục lục một dòng duy nhất.
    """
    if number is None:
        return False
    prev_number = _neighbor_lesson_number(lines, index, -1)
    next_number = _neighbor_lesson_number(lines, index, 1)
    return prev_number in (number - 1, number + 1) or next_number in (number - 1, number + 1)


def extract(text: str, query: str) -> tuple[str, str, list[str]]:
    lines = text.splitlines()
    found = headings(lines)
    if not found:
        raise SgkError(
            "File SGK không có dòng tiêu đề nào để cắt theo",
            "Kiểm lại file Markdown đã chuyển từ PDF; có thể bước chuyển đổi đã thất bại.",
        )
    target = _squash(normalise(query))
    matches = [item for item in found if _contains(_squash(normalise(item[1])), target)]
    if not matches:
        titles = [title for _, title in found]
        squashed_titles = [_squash(normalise(title)) for title in titles]
        nearest_squashed = difflib.get_close_matches(target, squashed_titles, n=3, cutoff=0)
        nearest = []
        for squashed in nearest_squashed:
            nearest.append(titles[squashed_titles.index(squashed)])
        raise SgkError(
            f"Không tìm thấy tiêu đề nào chứa {query!r}",
            f"Ba tiêu đề gần nhất trong file: {'; '.join(nearest)}. Sửa lại --bai cho khớp.",
        )

    def end_of(start_index: int, title: str) -> int:
        # Bài kết thúc ở tiêu đề "Bài <số>" kế tiếp có số bài khác, hoặc ở tiêu đề nông hơn hẳn
        # (ví dụ chương mới) — chỉ khi dòng đang cắt là tiêu đề Markdown. Đề mục con trong bài
        # (### I. ..., ## II. ...) không cắt bài. Khi dòng đang cắt là dòng thường ("Bài <số>"
        # không có '#', kiểu mục lục), ranh giới là dòng thường hoặc tiêu đề "Bài <số>" kế tiếp
        # có số bài LỚN HƠN — số bài tập đánh số trong bài (dòng thường) không cắt bài.
        start_level = _level(lines[start_index])
        start_number = _lesson_number(title)
        for index, other_title in found:
            if index <= start_index:
                continue
            level = _level(lines[index])
            if start_level:
                if level == 0:
                    continue
                other_number = _lesson_number(other_title)
                if (
                    _BAI_RE.match(normalise(other_title))
                    and other_number is not None
                    and other_number != start_number
                ):
                    return index
                if 0 < level < start_level:
                    return index
            else:
                other_number = _lesson_number(other_title)
                if (
                    other_number is not None
                    and start_number is not None
                    and other_number > start_number
                ):
                    return index
        return len(lines)

    # Dòng mục lục ("Bài 5. Ammonia ... 25") cũng là tiêu đề nhưng không có nội dung phía sau: bỏ qua.
    with_content = [
        item for item in matches
        if any(line.strip() for line in lines[item[0] + 1:end_of(item[0], item[1])])
    ]
    if with_content:
        md_headings = [item for item in with_content if _level(lines[item[0]]) > 0]
        if md_headings:
            start_index, heading = md_headings[0]
        else:
            # Một dòng mục lục "Bài <số>" cuối cùng của sách không có dòng thường nào sau nó
            # mang số bài lớn hơn, nên lát cắt của nó chạy đến hết file — trùm lên cả tiêu đề
            # bài thật kế tiếp. Dòng lặp lại đầu trang bên trong nội dung bài (PDF→text), hay
            # một bài thật mở đầu ngay bằng bài tập đánh số ("Bài 1.", "Bài 2."...), trông
            # giống hệt một tiêu đề "Bài <số>" nhưng dòng liền kề nó không mang đúng số liền
            # trước/sau — nên chỉ bỏ ứng viên "là mục lục" (_looks_like_toc_entry), không bỏ
            # theo việc lát cắt của nó trùm lên ứng viên khác. Còn lại rỗng thì quay về cách cũ:
            # lát cắt dài nhất trong mọi ứng viên có nội dung.
            non_list = [
                item for item in with_content
                if not _looks_like_toc_entry(lines, item[0], _lesson_number(item[1]))
            ]
            pool = non_list if non_list else with_content
            start_index, heading = max(
                pool,
                key=lambda item: len(
                    "\n".join(lines[item[0]:end_of(item[0], item[1])]).encode("utf-8")
                ),
            )
    else:
        start_index, heading = matches[0]
    body = "\n".join(lines[start_index:end_of(start_index, heading)]).strip() + "\n"
    warnings: list[str] = []
    size_kb = len(body.encode("utf-8")) / 1024
    if size_kb > MAX_KB:
        warnings.append(
            f"Phần cắt ra nặng {size_kb:.1f} KB, lớn hơn {MAX_KB} KB; có thể đã cắt sang bài sau."
        )
    if len(matches) > 1:
        warnings.append(
            f"Có {len(matches)} tiêu đề khớp {query!r}; đã lấy: {heading}"
        )
    return heading, body, warnings
