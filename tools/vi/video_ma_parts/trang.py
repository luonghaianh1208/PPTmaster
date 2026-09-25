"""Ghép trang HTML cho một cảnh: phong cách + khung + file cảnh + dữ liệu. Mọi thứ nhúng trong trang."""

from __future__ import annotations

import json
from pathlib import Path

from . import phong

RUNTIME = Path(__file__).resolve().parent / "runtime"
NGHIEM = Path(__file__).resolve().parents[1] / "thi_nghiem_parts" / "runtime"


def json_nhung(data) -> str:
    return json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")


def _doc(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def dung_trang(du: dict, model=None) -> str:
    scripts = []
    if du["loai"] == "thi-nghiem":
        scripts.append(_doc(NGHIEM / "khung.js"))
        scripts.append(model.js)
    for ten in ("khung-video.js", "hinh.js", "ban-tay.js", "may-quay.js"):
        scripts.append(_doc(RUNTIME / ten))
    scripts.append(_doc(RUNTIME / "canh" / f"{du['loai']}.js"))
    scripts.append(f"window.DU_CANH = {json_nhung(du)};\nTHI_VIDEO.khoiDong(window.DU_CANH);")
    body = "\n".join(f"<script>\n{s}\n</script>" for s in scripts)
    return ("<!doctype html>\n<html lang=\"vi\"><head><meta charset=\"utf-8\">"
            f"<style>\n{phong.font_css()}\n{_doc(RUNTIME / 'viet-tay.css')}\n</style></head>\n"
            f"<body><div id=\"khung\"></div>\n{body}\n</body></html>\n")
