"""Biểu tượng nét tabler-outline: chuẩn tên, đọc phần tử vẽ, gợi ý khi sai. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import difflib
import re
import xml.etree.ElementTree as ET
from pathlib import Path

THU_MUC = Path(__file__).resolve().parents[3] / "skills" / "ppt-master" / "templates" / "icons" / "tabler-outline"

NHAN_TOI_DA = 30
_TIEN_TO = "tabler-outline/"
_DUOI = ".svg"
_KHUNG_D = "M0 0h24v24H0z"
_TE_HINH = ("path", "circle", "rect", "line", "polyline", "polygon", "ellipse")
_SVG_NS = "{http://www.w3.org/2000/svg}"
_SO_GOI_Y = 5
_NGUONG_GOI_Y = 0.5
_MARKUP_RE = re.compile(r"\*\*|~|\^")


class HinhError(Exception):
    """Tên biểu tượng không có trong thư viện, hoặc dòng `hinh` sai định dạng `tên | nhãn`."""


def chuan_ten(s: str) -> str:
    ten = s.strip()
    if ten.lower().startswith(_TIEN_TO):
        ten = ten[len(_TIEN_TO):]
    if ten.lower().endswith(_DUOI):
        ten = ten[: -len(_DUOI)]
    return ten.strip().lower()


def _danh_sach() -> list:
    return sorted(p.stem for p in THU_MUC.glob("*.svg"))


def tach_minh_hoa(value: str) -> tuple:
    if "|" not in value:
        raise HinhError(f"dòng `hinh` phải có dạng `tên | nhãn`; thiếu dấu `|` trong `{value}`")
    ten, nhan = value.split("|", 1)
    ten, nhan = ten.strip(), nhan.strip()
    if not nhan:
        raise HinhError(f"dòng `hinh` phải có dạng `tên | nhãn`; thiếu nhãn trong `{value}`")
    so_ky_tu = len(_MARKUP_RE.sub("", nhan))
    if so_ky_tu > NHAN_TOI_DA:
        raise HinhError(f"nhãn hình `{nhan}` dài {so_ky_tu} ký tự, tối đa {NHAN_TOI_DA}")
    return ten, nhan


def doc(ten: str) -> dict:
    chuan = chuan_ten(ten)
    duong_dan = THU_MUC / f"{chuan}.svg"
    if not duong_dan.is_file():
        goi_y = difflib.get_close_matches(chuan, _danh_sach(), n=_SO_GOI_Y, cutoff=_NGUONG_GOI_Y)
        thong_bao = f"không có biểu tượng `{ten}` trong thư viện tabler-outline"
        if goi_y:
            thong_bao += ". Có thể bạn muốn: " + ", ".join(goi_y)
        raise HinhError(thong_bao)
    root = ET.fromstring(duong_dan.read_text(encoding="utf-8"))
    phan_tu = []
    for el in root:
        the = el.tag.removeprefix(_SVG_NS)
        if the not in _TE_HINH:
            continue
        thuoc_tinh = dict(el.attrib)
        if thuoc_tinh.get("stroke") == "none":
            continue
        if the == "path" and thuoc_tinh.get("d") == _KHUNG_D:
            continue
        phan_tu.append({"the": the, "thuocTinh": thuoc_tinh})
    return {"ten": chuan, "viewBox": root.attrib.get("viewBox", "0 0 24 24"), "phanTu": phan_tu}
