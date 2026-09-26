#!/usr/bin/env python3
"""Tìm nhạc nền giấy phép mở (CC0, CC BY) trên Openverse, tải bản mp3 và ghi nguồn cho video giải thích.

  python tools/vi/tim_nhac.py "<từ khoá>" -o <thư_mục_dự_án>/nhac [--so 3]

Chỉ lấy bản dài từ 60 giây. File lưu thành `<slug từ khoá>.mp3` (không ghi đè file đã có); nguồn thêm vào
`nguon.json` cùng thư mục, dạng {"items": [{filename, title, creator, license, license_url, source_url}]}.
stdout đúng một dòng JSON. Chỉ dùng thư viện chuẩn.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API = "https://api.openverse.org/v1/audio/"
GIAY_TOI_THIEU = 60
TOI_DA_BYTE = 40 * 1024 * 1024
HET_GIO = 30
TAC_NHAN = "PPTMaster-vi-tim-nhac/1.0"
GIAY_PHEP = {"cc0": "CC0", "by": "CC BY"}
FIX_MANG = "Kiểm tra kết nối Internet rồi chạy lại; Openverse bận thì thử lại sau vài phút."
FIX_INPUT = 'Dùng: python tools/vi/tim_nhac.py "<từ khoá>" -o <thư_mục_dự_án>/nhac [--so 3]'


class Loi(Exception):
    def __init__(self, step: str, message: str, fix: str) -> None:
        super().__init__(message)
        self.step, self.message, self.fix = step, message, fix


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def slug(chu: str) -> str:
    """Chữ thường không dấu nối bằng gạch ngang: "Nhạc nền Đà Lạt" -> "nhac-nen-da-lat"."""
    chu = unicodedata.normalize("NFD", chu.replace("đ", "d").replace("Đ", "D"))
    chu = "".join(c for c in chu if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", "-", chu).strip("-")[:60].strip("-")


def _mo(lay, url: str):
    req = urllib.request.Request(url, headers={"User-Agent": TAC_NHAN})
    try:
        return lay(req, timeout=HET_GIO)
    except (urllib.error.URLError, OSError, ValueError) as exc:
        raise Loi("mang", f"Không tải được {url}: {exc}", FIX_MANG) from exc


def _doc(tra_loi, gioi_han: int) -> bytes:
    try:
        with tra_loi as r:
            du_lieu = r.read(gioi_han + 1)
    except (urllib.error.URLError, OSError, ValueError) as exc:
        raise Loi("mang", f"Mất kết nối khi đang tải: {exc}", FIX_MANG) from exc
    return du_lieu


def _la_mp3(ban: dict) -> bool:
    kieu = (ban.get("filetype") or "").lower()
    if kieu:
        return kieu == "mp3"
    return urllib.parse.urlsplit(ban.get("url") or "").path.lower().endswith(".mp3")


def tim(tu_khoa: str, lay=urllib.request.urlopen) -> list:
    """Các bản CC0/CC BY, mp3, dài ≥ 60 giây của từ khoá, theo thứ tự Openverse trả về."""
    url = API + "?" + urllib.parse.urlencode({"q": tu_khoa, "license": "cc0,by", "page_size": 20})
    du_lieu = _doc(_mo(lay, url), 4 * 1024 * 1024)
    try:
        ket = json.loads(du_lieu.decode("utf-8"))
        ds = ket["results"]
        if not isinstance(ds, list):
            raise TypeError("results")
    except (ValueError, KeyError, TypeError) as exc:
        raise Loi("mang", f"Openverse trả về dữ liệu không đọc được ({type(exc).__name__}).", FIX_MANG) from exc
    ra = []
    for ban in ds:
        if not isinstance(ban, dict) or not ban.get("url"):
            continue
        if (ban.get("license") or "").lower() not in GIAY_PHEP:
            continue
        giay = ban.get("duration")
        if not isinstance(giay, (int, float)) or giay < GIAY_TOI_THIEU * 1000:
            continue
        if _la_mp3(ban):
            ra.append(ban)
    return ra


def _ten_giay_phep(ban: dict) -> str:
    ten = GIAY_PHEP[(ban.get("license") or "").lower()]
    phien_ban = str(ban.get("license_version") or "").strip()
    return f"{ten} {phien_ban}" if phien_ban else ten


def _la_file_mp3(du_lieu: bytes) -> bool:
    return du_lieu.startswith(b"ID3") or (len(du_lieu) > 1 and du_lieu[0] == 0xFF and du_lieu[1] & 0xE0 == 0xE0)


def _doc_nguon(tep: Path) -> dict:
    if not tep.is_file():
        return {"items": []}
    sai = f"`{tep.name}` sai cấu trúc; sửa lại theo dạng {{\"items\": [...]}} hoặc xoá file rồi chạy lại."
    try:
        du_lieu = json.loads(tep.read_text(encoding="utf-8-sig"))
    except ValueError as exc:
        raise Loi("input", f"`{tep}` không phải JSON hợp lệ ({exc}).", sai) from exc
    except OSError as exc:
        raise Loi("write", f"Không đọc được `{tep}`: {exc}", "Đóng file đang mở rồi chạy lại.") from exc
    if not isinstance(du_lieu, dict) or not isinstance(du_lieu.get("items", []), list):
        raise Loi("input", f"`{tep}` sai cấu trúc.", sai)
    du_lieu.setdefault("items", [])
    return du_lieu


def _ten_trong(thu_muc: Path, goc: str, da_dung: set) -> str:
    k = 1
    while True:
        ten = f"{goc}.mp3" if k == 1 else f"{goc}-{k}.mp3"
        if ten not in da_dung and not (thu_muc / ten).exists():
            return ten
        k += 1


def chay(tu_khoa: str, thu_muc: Path, so: int, lay, warnings: list) -> list:
    goc = slug(tu_khoa)
    if not goc:
        raise Loi("input", "Từ khoá cần có ít nhất một chữ cái hoặc chữ số.", FIX_INPUT)
    try:
        thu_muc.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise Loi("write", f"Không tạo được thư mục `{thu_muc}`: {exc}", "Chọn thư mục khác cho -o rồi chạy lại.") from exc
    tep_nguon = thu_muc / "nguon.json"
    nguon = _doc_nguon(tep_nguon)
    cac_ban = tim(tu_khoa, lay=lay)
    if not cac_ban:
        raise Loi("input", f"Openverse không có bản CC0/CC BY dạng mp3 dài từ {GIAY_TOI_THIEU} giây cho `{tu_khoa}`.",
                  "Thử từ khoá khác, ngắn hơn, bằng tiếng Anh (ví dụ \"calm piano\").")
    files: list = []
    moi: list = []
    for ban in cac_ban:
        if len(files) >= so:
            break
        du_lieu = _doc(_mo(lay, ban["url"]), TOI_DA_BYTE)
        if len(du_lieu) > TOI_DA_BYTE or not _la_file_mp3(du_lieu):
            warnings.append(f"Bỏ qua `{ban.get('title') or ban['url']}`: file tải về không phải mp3 hoặc quá lớn.")
            continue
        ten = _ten_trong(thu_muc, goc, set(files))
        tam = thu_muc / (ten + ".tam")
        try:
            tam.write_bytes(du_lieu)
            os.replace(tam, thu_muc / ten)
        except OSError as exc:
            raise Loi("write", f"Không ghi được `{thu_muc / ten}`: {exc}", "Kiểm tra ổ đĩa rồi chạy lại.") from exc
        files.append(ten)
        moi.append({"filename": ten, "title": str(ban.get("title") or "").strip(),
                    "creator": str(ban.get("creator") or "").strip(), "license": _ten_giay_phep(ban),
                    "license_url": ban.get("license_url") or "", "source_url": ban.get("foreign_landing_url") or ""})
    if not files:
        raise Loi("input", f"Không tải được bản mp3 hợp lệ nào cho `{tu_khoa}`.", "Thử từ khoá khác rồi chạy lại.")
    if len(files) < so:
        warnings.append(f"Chỉ tìm được {len(files)} bản phù hợp (yêu cầu {so}).")
    nguon["items"] = [i for i in nguon["items"] if not (isinstance(i, dict) and i.get("filename") in files)] + moi
    try:
        tam = thu_muc / "nguon.json.tam"
        tam.write_text(json.dumps(nguon, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        os.replace(tam, tep_nguon)
    except OSError as exc:
        raise Loi("write", f"Không ghi được `{tep_nguon}`: {exc}", "Đóng file đang mở rồi chạy lại.") from exc
    return moi


def main(argv=None, lay=urllib.request.urlopen) -> int:
    base = {"ready": False, "files": [], "ban": [], "thu_muc": None}
    ap = argparse.ArgumentParser(description="Tìm nhạc nền CC0/CC BY trên Openverse", add_help=False)
    ap.add_argument("tu_khoa")
    ap.add_argument("-o", dest="thu_muc", required=True)
    ap.add_argument("--so", type=int, default=1)
    warnings: list = []
    try:
        try:
            with contextlib.redirect_stderr(io.StringIO()):
                args = ap.parse_args(argv)
        except SystemExit as exc:
            raise Loi("input", "Sai tham số dòng lệnh.", FIX_INPUT) from exc
        if not 1 <= args.so <= 5:
            raise Loi("input", "`--so` là số bản cần tải, từ 1 đến 5.", FIX_INPUT)
        thu_muc = Path(args.thu_muc).resolve()
        moi = chay(args.tu_khoa.strip(), thu_muc, args.so, lay, warnings)
        emit({**base, "ready": True, "files": [m["filename"] for m in moi], "ban": moi, "thu_muc": str(thu_muc),
              "warnings": warnings, "error": None})
        return 0
    except Loi as exc:
        error = {"step": exc.step, "message": exc.message, "fix": exc.fix}
    except OSError as exc:
        error = {"step": "write", "message": f"Không ghi được file: {exc}", "fix": "Kiểm tra ổ đĩa rồi chạy lại."}
    emit({**base, "warnings": warnings, "error": error})
    return 1


if __name__ == "__main__":
    sys.exit(main())
