"""Ảnh thật trong anh/: đọc kích thước, nhúng data:, dựng dòng nguồn. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import base64
import json
import struct
from pathlib import Path

DINH_DANG = (".jpg", ".jpeg", ".png", ".webp")
TOI_DA = 8 * 1024 * 1024
_MIME = {".jpg": "jpeg", ".jpeg": "jpeg", ".png": "png", ".webp": "webp"}
_SOF_MARKERS = (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF)


class AnhError(Exception):
    """Ảnh trong anh/ không đọc được, sai định dạng, quá lớn, hoặc chưa có nguồn."""


def _kich_thuoc_png(du_lieu: bytes) -> tuple:
    rong, cao = struct.unpack(">II", du_lieu[16:24])
    return rong, cao


def _kich_thuoc_jpeg(du_lieu: bytes) -> tuple:
    i = 2
    n = len(du_lieu)
    while i + 3 < n:
        if du_lieu[i] != 0xFF:
            i += 1
            continue
        marker = du_lieu[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        if marker == 0xD9:
            break
        do_dai = struct.unpack(">H", du_lieu[i + 2:i + 4])[0]
        if marker in _SOF_MARKERS:
            cao, rong = struct.unpack(">HH", du_lieu[i + 5:i + 9])
            return rong, cao
        i += 2 + do_dai
    raise struct.error("không thấy đoạn SOF")


def _kich_thuoc_webp(du_lieu: bytes) -> tuple:
    fourcc = du_lieu[12:16]
    if fourcc == b"VP8X":
        rong = int.from_bytes(du_lieu[24:27], "little") + 1
        cao = int.from_bytes(du_lieu[27:30], "little") + 1
        return rong, cao
    if fourcc == b"VP8L":
        bits = int.from_bytes(du_lieu[21:25], "little")
        rong = (bits & 0x3FFF) + 1
        cao = ((bits >> 14) & 0x3FFF) + 1
        return rong, cao
    if fourcc == b"VP8 ":
        rong = struct.unpack("<H", du_lieu[26:28])[0] & 0x3FFF
        cao = struct.unpack("<H", du_lieu[28:30])[0] & 0x3FFF
        return rong, cao
    raise struct.error(f"khối WEBP lạ: {fourcc!r}")


def _kich_thuoc(duoi: str, du_lieu: bytes, ten_file: str) -> tuple:
    try:
        if duoi == ".png":
            return _kich_thuoc_png(du_lieu)
        if duoi in (".jpg", ".jpeg"):
            return _kich_thuoc_jpeg(du_lieu)
        return _kich_thuoc_webp(du_lieu)
    except (struct.error, IndexError) as exc:
        raise AnhError(f"không đọc được kích thước `anh/{ten_file}` (file hỏng?)") from exc


def _nguon_tu_manifest(thu_muc_du_an: Path, ten_file: str) -> str:
    manifest = Path(thu_muc_du_an) / "anh" / "image_sources.json"
    if not manifest.is_file():
        return ""
    try:
        du_lieu = json.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return ""
    for muc in du_lieu.get("items", []):
        if muc.get("filename") == ten_file:
            phan = [muc.get("author") or "", muc.get("license_name") or muc.get("license") or "", muc.get("provider") or ""]
            phan = [p for p in phan if p]
            if not phan:
                return ""
            return "Ảnh: " + " · ".join(phan)
    return ""


def doc(thu_muc_du_an: Path, ten_file: str, nguon_tay: str | None) -> dict:
    duoi = Path(ten_file).suffix.lower()
    if duoi not in DINH_DANG:
        raise AnhError(f"`{ten_file}` không phải định dạng ảnh cho phép ({', '.join(DINH_DANG)})")
    duong_dan = Path(thu_muc_du_an) / "anh" / ten_file
    if not duong_dan.is_file():
        raise AnhError(f"không có file `anh/{ten_file}`")
    kich_thuoc_byte = duong_dan.stat().st_size
    if kich_thuoc_byte > TOI_DA:
        raise AnhError(f"`anh/{ten_file}` nặng {kich_thuoc_byte} byte, tối đa {TOI_DA}")
    du_lieu = duong_dan.read_bytes()
    rong, cao = _kich_thuoc(duoi, du_lieu, ten_file)
    nguon = nguon_tay.strip() if nguon_tay else _nguon_tu_manifest(thu_muc_du_an, ten_file)
    if not nguon:
        raise AnhError(f"`anh/{ten_file}` chưa có nguồn: ghi `nguon:` trong cảnh, hoặc thêm bản ghi cho "
                        f"`{ten_file}` vào `anh/image_sources.json`")
    data_url = f"data:image/{_MIME[duoi]};base64," + base64.b64encode(du_lieu).decode("ascii")
    return {"dataUrl": data_url, "nguon": nguon, "rong": rong, "cao": cao}
