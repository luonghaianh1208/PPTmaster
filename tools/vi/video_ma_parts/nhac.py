"""Nhạc nền trong nhac/: chặn đường dẫn, đọc nguồn, đo thời lượng bằng ffprobe. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import json
import subprocess
import unicodedata
from pathlib import Path

from .anh import _bo_dia_chi_web, _hop_le, _tim_theo_nfc

DINH_DANG = (".mp3", ".m4a", ".wav", ".ogg")


class NhacError(Exception):
    """File nhạc nền không hợp lệ, không đọc được, hoặc chưa có nguồn."""


def _nfc(chu) -> str:
    return unicodedata.normalize("NFC", str(chu or ""))


def _nguon_tu_manifest(thu_muc_du_an: Path, ten_file: str) -> str:
    manifest = Path(thu_muc_du_an) / "nhac" / "nguon.json"
    if not manifest.is_file():
        return ""
    sai = '`nhac/nguon.json` sai cấu trúc: cần dạng {"items": [{"filename": ..., "title": ..., "creator": ..., "license": ...}]}'
    try:
        du_lieu = json.loads(manifest.read_text(encoding="utf-8-sig"))
    except OSError as exc:
        raise NhacError(f"không đọc được `nhac/nguon.json`: {exc}") from exc
    except ValueError as exc:
        raise NhacError(f"`nhac/nguon.json` không phải JSON hợp lệ ({exc})") from exc
    if not isinstance(du_lieu, dict) or not isinstance(du_lieu.get("items", []), list):
        raise NhacError(sai)
    cac_muc = du_lieu.get("items", [])
    if not all(isinstance(muc, dict) for muc in cac_muc):
        raise NhacError(sai + "; mỗi phần tử của `items` phải là một bản ghi {...}")
    for muc in cac_muc:
        if _nfc(muc.get("filename")) == _nfc(ten_file):
            phan = [p for p in (_bo_dia_chi_web(muc.get(k)) for k in ("title", "creator", "license")) if p]
            return "Nhạc: " + " · ".join(phan) if phan else ""
    return ""


def thoi_luong(duong_dan: Path, run=subprocess.run):
    """Thời lượng (giây) theo ffprobe; None khi máy chưa có ffprobe (bước dựng tự kiểm FFmpeg sau)."""
    cmd = ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1",
           str(duong_dan)]
    try:
        proc = run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
    except OSError:
        return None
    except subprocess.TimeoutExpired as exc:
        raise NhacError(f"ffprobe đọc `nhac/{duong_dan.name}` quá lâu (file hỏng?)") from exc
    try:
        giay = float((proc.stdout or "").strip()) if proc.returncode == 0 else 0.0
    except ValueError:
        giay = 0.0
    if not giay > 0:
        raise NhacError(f"không đọc được `nhac/{duong_dan.name}` (file hỏng hoặc không phải file âm thanh)")
    return giay


def doc(thu_muc_du_an: Path, ten_file: str, nguon_tay: str | None, run=subprocess.run) -> dict:
    loi_ten = f"`{ten_file}` không hợp lệ: `nhac-nen` chỉ được là tên file nằm trong `nhac/`, không phải đường dẫn."
    if not _hop_le(ten_file):
        raise NhacError(loi_ten)
    if Path(ten_file).suffix.lower() not in DINH_DANG:
        raise NhacError(f"`nhac/{ten_file}` không phải định dạng nhạc cho phép ({', '.join(DINH_DANG)})")
    thu_muc_nhac = (Path(thu_muc_du_an) / "nhac").resolve()
    duong_dan = (thu_muc_nhac / ten_file).resolve()
    if not duong_dan.is_relative_to(thu_muc_nhac):
        raise NhacError(loi_ten)
    if not duong_dan.is_file():
        duong_dan = _tim_theo_nfc(thu_muc_nhac, ten_file) or duong_dan
    if not duong_dan.is_file():
        raise NhacError(f"không có file `nhac/{ten_file}`")
    if nguon_tay and nguon_tay.strip():
        nguon = nguon_tay.strip()
        nguon = nguon if nguon.lower().startswith("nhạc") else "Nhạc: " + nguon
    else:
        nguon = _nguon_tu_manifest(thu_muc_du_an, ten_file)
    if not nguon:
        raise NhacError(f"`nhac/{ten_file}` chưa có nguồn: ghi khoá đầu `nguon-nhac:`, hoặc thêm bản ghi cho "
                        f"`{ten_file}` vào `nhac/nguon.json` (công cụ tools/vi/tim_nhac.py tự ghi)")
    return {"duong_dan": duong_dan, "nguon": nguon, "giay": thoi_luong(duong_dan, run=run)}
