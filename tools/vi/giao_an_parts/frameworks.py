"""Đọc khung năng lực số và năng lực AI từ file tài liệu, và kiểm mã của một giáo án.

Nguồn mã duy nhất là docs/vi/tro-ly/nang-luc-so-va-ai.md; module này không khai lại mã nào.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
DOC_PATH = REPO_ROOT / "docs" / "vi" / "tro-ly" / "nang-luc-so-va-ai.md"
NLS_HEADING = "## Khung năng lực số (mọi môn)"
AI_HEADING = "## Khung năng lực AI theo môn"
NO_FRAMEWORK = "(chưa có khung mã cho môn này)"
NO_CODE = "(chưa có mã)"
SPECIAL_TRACK = "hệ chuyên"
TRACK_ALIASES = ("chuyen", "he chuyen")
FIX_RELOAD = "Tải lại bản đầy đủ của bộ công cụ; đừng sửa khuôn của nang-luc-so-va-ai.md."

_CODE_RE = re.compile(r"^-\s+`([^`]+)`")
_FRAME_RE = re.compile(r"^###\s+(.+?)\s+—\s+(.+?)\s+—\s+khung\s+`([^`]+)`\s*$")


class FrameworkError(Exception):
    """Mã năng lực sai, hoặc không đọc được bảng mã."""

    def __init__(self, message: str, fix: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.fix = fix


def normalise(text: str) -> str:
    """Bỏ dấu, hạ chữ thường, gộp khoảng trắng — để so tên môn không lệ thuộc cách viết."""
    decomposed = unicodedata.normalize("NFD", text)
    without_marks = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return " ".join(without_marks.lower().split())


@dataclass
class Frameworks:
    nls: list[str] = field(default_factory=list)
    ai: dict[str, list[str]] = field(default_factory=dict)
    index: dict[tuple[str, str], str] = field(default_factory=dict)
    labels: dict[str, list[str]] = field(default_factory=dict)

    def framework_for(self, subject: str, grade: str, track: str = "") -> str | None:
        scope = SPECIAL_TRACK if normalise(track) in TRACK_ALIASES else f"lớp {grade}"
        return self.index.get((normalise(subject), normalise(scope)))

    def ai_codes_for(self, subject: str, grade: str, track: str = "") -> list[str]:
        name = self.framework_for(subject, grade, track)
        return list(self.ai.get(name, [])) if name else []


def load(text: str) -> Frameworks:
    frameworks = Frameworks()
    section = ""
    current = ""
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("## "):
            section = line
            current = ""
            continue
        if section == NLS_HEADING:
            match = _CODE_RE.match(line)
            if match:
                frameworks.nls.append(match.group(1))
            continue
        if section == AI_HEADING:
            frame = _FRAME_RE.match(line)
            if frame:
                subject, scope, name = frame.groups()
                current = name
                frameworks.ai.setdefault(name, [])
                frameworks.index[(normalise(subject), normalise(scope))] = name
                scopes = frameworks.labels.setdefault(subject, [])
                if scope not in scopes:
                    scopes.append(scope)
                continue
            match = _CODE_RE.match(line)
            if match and current:
                frameworks.ai[current].append(match.group(1))
    if not frameworks.nls:
        raise FrameworkError(f"Không đọc được mã năng lực số nào trong {DOC_PATH.name}", FIX_RELOAD)
    if not frameworks.ai:
        raise FrameworkError(f"Không đọc được khung năng lực AI nào trong {DOC_PATH.name}", FIX_RELOAD)
    return frameworks


def load_file(path: Path = DOC_PATH) -> Frameworks:
    try:
        return load(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise FrameworkError(f"Không đọc được {path}: {exc}", FIX_RELOAD) from None


def validate(lesson, frameworks: Frameworks) -> None:
    """Kiểm mã NLS và mã AI của một giáo án. Lỗi thì nêu rõ mã sai và cách sửa."""
    unknown_nls = [code for code in lesson.nls_codes if code not in frameworks.nls]
    if unknown_nls:
        raise FrameworkError(
            "Mã năng lực số không có trong bảng của Bộ: " + ", ".join(unknown_nls),
            "Dùng một trong các mã: " + ", ".join(frameworks.nls),
        )

    subject = lesson.meta["subject"]
    grade = lesson.meta["grade"]
    track = lesson.meta.get("track", "")
    name = frameworks.framework_for(subject, grade, track)
    allowed = frameworks.ai_codes_for(subject, grade, track)

    if name is None:
        if not lesson.ai_no_framework:
            labels_text = "; ".join(
                f"{label_subject} ({', '.join(scopes)})"
                for label_subject, scopes in frameworks.labels.items()
            )
            raise FrameworkError(
                f"Môn {subject} lớp {grade} chưa có khung mã AI trong {DOC_PATH.name}. "
                f"Khung AI hiện có: {labels_text}.",
                "Nếu môn này có trong danh sách trên thì sửa dòng subject: cho đúng tên; "
                f"nếu không, ghi dòng đầu của mục '### Nang luc AI' là '- {NO_FRAMEWORK}', "
                f"và dùng 'ai: {NO_CODE}' cho hoạt động AI.",
            )
        coded = [
            line.split(" — ", 1)[0]
            for line in lesson.ai_lines[1:]
            if re.match(r"^AI-\S+\s+—\s+", line)
        ]
        if coded:
            raise FrameworkError(
                "Môn này chưa có khung mã AI nhưng mục tiêu vẫn ghi mã: " + ", ".join(coded),
                f"Bỏ mã ở đầu các dòng đó; dòng đầu là '- {NO_FRAMEWORK}', các dòng sau viết bằng lời.",
            )
        stray = [code for code in _activity_codes(lesson, "ai") if code != NO_CODE]
        if stray:
            raise FrameworkError(
                "Môn này chưa có khung mã AI nhưng hoạt động vẫn ghi mã: " + ", ".join(stray),
                f"Bỏ các mã đó; dùng 'ai: {NO_CODE}'.",
            )
    else:
        if lesson.ai_no_framework:
            raise FrameworkError(
                f"Môn {subject} lớp {grade} đã có khung {name} nên không được ghi "
                f"'{NO_FRAMEWORK}'",
                "Dùng một trong các mã: " + ", ".join(allowed),
            )
        unknown_ai = [code for code in lesson.ai_codes if code not in allowed]
        if unknown_ai:
            raise FrameworkError(
                f"Mã AI không thuộc khung {name} của môn {subject} lớp {grade}: "
                + ", ".join(unknown_ai),
                "Dùng một trong các mã: " + ", ".join(allowed),
            )
        if NO_CODE in _activity_codes(lesson, "ai"):
            raise FrameworkError(
                f"Môn {subject} lớp {grade} đã có khung {name} nên hoạt động không được ghi "
                f"'ai: {NO_CODE}'",
                "Dùng một trong các mã: " + ", ".join(allowed),
            )

    for activity in lesson.activities:
        for code in activity.nls:
            if code not in frameworks.nls:
                raise FrameworkError(
                    f"Mã năng lực số không có trong bảng của Bộ: {code} (hoạt động {activity.title!r})",
                    "Dùng một trong các mã: " + ", ".join(frameworks.nls),
                )
        if name is not None:
            for code in activity.ai:
                if code != NO_CODE and code not in allowed:
                    raise FrameworkError(
                        f"Mã AI không thuộc khung {name} của môn {subject} lớp {grade}: {code} "
                        f"(hoạt động {activity.title!r})",
                        "Dùng một trong các mã: " + ", ".join(allowed),
                    )

    for activity in lesson.activities:
        for code in activity.nls:
            if code not in lesson.nls_codes:
                raise FrameworkError(
                    f"Hoạt động {activity.title!r} dùng mã {code} chưa khai ở mục "
                    "'### Nang luc so'",
                    "Thêm mã đó vào mục tiêu, hoặc bỏ khỏi hoạt động.",
                )
        for code in activity.ai:
            if code == NO_CODE:
                continue
            if code not in lesson.ai_codes:
                raise FrameworkError(
                    f"Hoạt động {activity.title!r} dùng mã {code} chưa khai ở mục "
                    "'### Nang luc AI'",
                    "Thêm mã đó vào mục tiêu, hoặc bỏ khỏi hoạt động.",
                )


def _activity_codes(lesson, attribute: str) -> list[str]:
    codes: list[str] = []
    for activity in lesson.activities:
        for code in getattr(activity, attribute):
            if code not in codes:
                codes.append(code)
    return codes
