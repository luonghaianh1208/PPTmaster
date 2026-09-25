"""Đọc video.md thành dữ liệu; lỗi luôn kèm số dòng. Ngữ pháp ở docs/vi/tro-ly/video-giai-thich.md. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re
from dataclasses import dataclass

META_REQUIRED = ("tieu-de", "mon", "lop")
META_CHOICES = {
    "phong-cach": ("viet-tay",),
    "giong": ("nu", "nam"),
    "toc-do": ("cham", "vua", "nhanh"),
    "phu-de": ("hinh", "file", "khong"),
    "ban-tay": ("co", "khong"),
    "may-quay": ("co", "khong"),
    "chuyen-canh": ("lau-bang", "khong"),
}
META_DEFAULTS = {
    "phong-cach": "viet-tay", "giong": "nu", "toc-do": "vua", "phu-de": "hinh",
    "ban-tay": "co", "may-quay": "co", "chuyen-canh": "lau-bang",
}

# loại cảnh -> (trường đơn bắt buộc, trường đơn tuỳ chọn, trường lặp {khoá: (tối thiểu, tối đa)})
SCENE_SPEC = {
    "tieu-de": (("chu",), ("phu", "hinh", "anh", "nguon"), {}),
    "khai-niem": (("thuat-ngu", "dinh-nghia"), ("hinh", "anh", "nguon"), {}),
    "cong-thuc": (("bieu-thuc",), ("hinh", "anh", "nguon"), {"giai-thich": (0, 4)}),
    "y-tung-y": (("tieu-de",), ("hinh", "anh", "nguon"), {"y": (1, 6)}),
    "quy-trinh": (("tieu-de",), (), {"buoc": (2, 5)}),
    "so-sanh": (("tieu-de", "trai", "phai"), (), {"y-trai": (1, 4), "y-phai": (1, 4)}),
    "do-thi": (("tieu-de", "truc-ngang", "truc-doc"), (), {"diem": (2, 12)}),
    "thi-nghiem": (("mau",), ("do",), {"tham-so": (0, 99)}),
    "minh-hoa": (("tieu-de",), (), {"hinh": (1, 3)}),
    "anh": (("anh", "chu-thich"), ("nguon",), {}),
}
SCENE_TYPES = tuple(SCENE_SPEC)

_KEY_RE = re.compile(r"^([a-z][a-z0-9-]*):\s*(.*)$")
_SCENE_RE = re.compile(r"^##\s+Cảnh\s+(\d+)\s*$")
_URL_RE = re.compile(r"https?://|www\.", re.IGNORECASE)
_POINT_RE = re.compile(r"^-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?$")
_PARAM_RE = re.compile(r"^\d+(?:\.\d+)?\s+[a-z0-9-]+\s+-?\d+(?:\.\d+)?$")


class ParseError(Exception):
    """Lỗi trong video.md; luôn kèm số dòng để AI sửa được đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Scene:
    so: int
    dong: int
    loai: str
    loi: str
    truong: dict
    dong_truong: dict


@dataclass
class Video:
    meta: dict
    canh: list


def _check_value(no: int, key: str, value: str) -> None:
    if value == "":
        raise ParseError(no, f"`{key}` đang trống.")
    if _URL_RE.search(value):
        raise ParseError(no, f"`{key}` chứa địa chỉ web. Không chèn địa chỉ web vào video.")


def _read_meta(lines: list, start: int) -> tuple:
    meta: dict = {}
    i = start
    while i < len(lines):
        raw = lines[i].strip()
        if raw == "---":
            break
        if raw:
            match = _KEY_RE.match(raw)
            if match is None:
                raise ParseError(i + 1, "Dòng thông tin phải có dạng `khoá: giá trị`.")
            key, value = match.group(1), match.group(2).strip()
            if key not in META_REQUIRED and key not in META_CHOICES:
                raise ParseError(i + 1, f"Khoá `{key}` không có trong khối thông tin.")
            if key in meta:
                raise ParseError(i + 1, f"Khoá `{key}` bị lặp.")
            _check_value(i + 1, key, value)
            if key in META_CHOICES and value not in META_CHOICES[key]:
                raise ParseError(i + 1, f"`{key}` phải là một trong: {', '.join(META_CHOICES[key])}.")
            meta[key] = value
        i += 1
    else:
        raise ParseError(start, "Khối thông tin chưa đóng bằng dòng `---`.")
    for key in META_REQUIRED:
        if key not in meta:
            raise ParseError(i + 1, f"Khối thông tin thiếu `{key}`.")
    for key, default in META_DEFAULTS.items():
        meta.setdefault(key, default)
    return meta, i + 1


def _finish(so: int, dong0: int, fields: list) -> Scene:
    truong: dict = {}
    dong_truong: dict = {}
    for key, value, no in fields:
        truong.setdefault(key, []).append(value)
        dong_truong.setdefault(key, []).append(no)
    for key in ("loai", "loi"):
        if key not in truong:
            raise ParseError(dong0, f"Cảnh {so} thiếu `{key}`.")
        if len(truong[key]) > 1:
            raise ParseError(dong_truong[key][1], f"`{key}` bị lặp trong Cảnh {so}.")
    loai = truong["loai"][0]
    if loai not in SCENE_SPEC:
        raise ParseError(dong_truong["loai"][0], f"`loai` phải là một trong: {', '.join(SCENE_TYPES)}.")
    required, optional, repeated = SCENE_SPEC[loai]
    allowed = {"loai", "loi", *required, *optional, *repeated}
    for key, value, no in fields:
        if key not in allowed:
            raise ParseError(no, f"Cảnh loại `{loai}` không có trường `{key}`.")
    for key in (*required, *optional):
        if key in truong and len(truong[key]) > 1:
            raise ParseError(dong_truong[key][1], f"`{key}` bị lặp trong Cảnh {so}.")
    for key in required:
        if key not in truong:
            raise ParseError(dong0, f"Cảnh {so} (loại `{loai}`) thiếu `{key}`.")
    for key, (low, high) in repeated.items():
        count = len(truong.get(key, []))
        if count < low:
            raise ParseError(dong0, f"Cảnh {so} (loại `{loai}`) cần ít nhất {low} dòng `{key}`.")
        if count > high:
            raise ParseError(dong_truong[key][high], f"Cảnh loại `{loai}` chỉ có tối đa {high} dòng `{key}`.")
    if "hinh" in truong and "anh" in truong:
        dong_sau = max(dong_truong["hinh"][0], dong_truong["anh"][0])
        raise ParseError(dong_sau, f"Cảnh {so} chỉ được có `hinh` hoặc `anh`, không cả hai.")
    if "nguon" in truong and "anh" not in truong:
        raise ParseError(dong_truong["nguon"][0],
                         f"`nguon` chỉ dùng kèm `anh` (dòng nguồn của ảnh thật); Cảnh {so} chưa có `anh`.")
    for value, no in zip(truong.get("diem", []), dong_truong.get("diem", [])):
        if _POINT_RE.match(value) is None:
            raise ParseError(no, "Điểm đồ thị phải có dạng `x, y` (hai số, dấu thập phân là dấu chấm).")
    for value, no in zip(truong.get("tham-so", []), dong_truong.get("tham-so", [])):
        if _PARAM_RE.match(value) is None:
            raise ParseError(no, "Dòng `tham-so` phải có dạng `<giây> <mã> <giá trị>`, ví dụ `0 chieu-dai 0.4`.")
    loi = truong.pop("loi")[0]
    truong.pop("loai")
    dong_truong.pop("loai")
    dong_truong.pop("loi")
    return Scene(so=so, dong=dong0, loai=loai, loi=loi, truong=truong, dong_truong=dong_truong)


def parse(text: str) -> Video:
    lines = text.lstrip("﻿").splitlines()
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i >= len(lines) or lines[i].strip() != "---":
        raise ParseError(i + 1, "video.md phải mở đầu bằng khối thông tin giữa hai dòng `---`.")
    meta, i = _read_meta(lines, i + 1)
    scenes: list = []
    current = None
    for index in range(i, len(lines)):
        no = index + 1
        raw = lines[index].strip()
        if not raw:
            continue
        heading = _SCENE_RE.match(raw)
        if heading:
            if current is not None:
                scenes.append(_finish(*current))
            expected = len(scenes) + 1
            if int(heading.group(1)) != expected:
                raise ParseError(no, f"Cảnh phải đánh số liên tiếp; ở đây phải là `## Cảnh {expected}`.")
            current = (expected, no, [])
            continue
        if current is None:
            raise ParseError(no, "Sau khối thông tin phải là `## Cảnh 1`.")
        match = _KEY_RE.match(raw)
        if match is None:
            raise ParseError(no, "Dòng trong cảnh phải có dạng `khoá: giá trị`.")
        key, value = match.group(1), match.group(2).strip()
        _check_value(no, key, value)
        current[2].append((key, value, no))
    if current is not None:
        scenes.append(_finish(*current))
    if not scenes:
        raise ParseError(i + 1, "video.md chưa có cảnh nào; bắt đầu bằng `## Cảnh 1`.")
    return Video(meta=meta, canh=scenes)
