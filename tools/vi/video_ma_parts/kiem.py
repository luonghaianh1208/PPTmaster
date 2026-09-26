"""Kiểm giới hạn chữ và nội dung cảnh của video.md. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re
from pathlib import Path

from thi_nghiem_parts import thu_vien

from . import anh, hinh
from .parse import ParseError, Scene, Video

LIMITS = {
    ("tieu-de", "chu"): 90, ("tieu-de", "phu"): 90,
    ("khai-niem", "thuat-ngu"): 60, ("khai-niem", "dinh-nghia"): 220,
    ("cong-thuc", "bieu-thuc"): 90, ("cong-thuc", "giai-thich"): 60,
    ("y-tung-y", "tieu-de"): 90, ("y-tung-y", "y"): 60,
    ("quy-trinh", "tieu-de"): 90, ("quy-trinh", "buoc"): 50,
    ("so-sanh", "tieu-de"): 90, ("so-sanh", "trai"): 24, ("so-sanh", "phai"): 24,
    ("so-sanh", "y-trai"): 60, ("so-sanh", "y-phai"): 60,
    ("do-thi", "tieu-de"): 90, ("do-thi", "truc-ngang"): 40, ("do-thi", "truc-doc"): 40,
    ("minh-hoa", "tieu-de"): 90, ("anh", "chu-thich"): 90,
}
LOI_DAI = 700
MAX_THAM_SO = 3
MAX_DO = 3
_MARKUP_RE = re.compile(r"\*\*|~|\^")
# Cùng ngữ pháp với catDanhDau trong runtime/khung-video.js.
_CUM_RE = re.compile(r"==(.+?)==|\(\((.+?)\)\)|__(.+?)__")
_SO_RE = re.compile(r"\{\{(.*?)\}\}")
_SO_DUNG = re.compile(r"-?\d+(?:\.\d+)?")
MAX_CUM = 3


class CanhError(Exception):
    """Nội dung cảnh sai; luôn nêu số cảnh."""

    def __init__(self, so: int, message: str) -> None:
        super().__init__(f"Cảnh {so}: {message}")
        self.so = so
        self.message = message


def _trong_cum(m: re.Match) -> str:
    return next(g for g in m.groups() if g is not None)


def hien_thi(chu: str) -> int:
    chu = _CUM_RE.sub(_trong_cum, chu)
    chu = _SO_RE.sub(lambda m: m.group(1).replace(".", ","), chu)
    return len(_MARKUP_RE.sub("", chu))


def kiem_danh_dau(key: str, value: str, no: int) -> None:
    """Cụm nhấn `==`, `((…))`, `__` và số chạy `{{…}}`: không lồng, có đóng, tối đa 3 cụm, số dùng dấu chấm."""
    cac_cum = list(_CUM_RE.finditer(value))
    for m in cac_cum:
        trong = _trong_cum(m)
        if _CUM_RE.search(trong) or any(dau in trong for dau in ("==", "((", "__")):
            raise ParseError(no, f"`{key}` có cụm nhấn lồng trong cụm khác (`{m.group(0)}`). Mỗi cụm nhấn đứng riêng.")
    con_lai = _CUM_RE.sub("", value)
    for dau, dong in (("==", "=="), ("((", "))"), ("__", "__")):
        if dau in con_lai:
            raise ParseError(no, f"`{key}` có `{dau}` chưa đóng bằng `{dong}`.")
    if len(cac_cum) > MAX_CUM:
        raise ParseError(no, f"`{key}` có {len(cac_cum)} cụm nhấn, tối đa {MAX_CUM} cụm mỗi dòng.")
    for m in _SO_RE.finditer(value):
        if _SO_DUNG.fullmatch(m.group(1)) is None:
            raise ParseError(no, f"`{key}`: `{m.group(0)}` phải là một số, dấu thập phân là dấu chấm (ví dụ `{{{{1500.5}}}}`).")
    con_lai = _SO_RE.sub("", value)
    if "{{" in con_lai or "}}" in con_lai:
        raise ParseError(no, f"`{key}` có `{{{{` hoặc `}}}}` chưa thành cặp; số chạy viết dạng `{{{{12}}}}`.")


def tham_so_theo_thoi_gian(scene: Scene) -> dict:
    out: dict = {}
    for value in scene.truong.get("tham-so", []):
        giay, ma, gia_tri = value.split()
        out.setdefault(ma, []).append((float(giay), float(gia_tri)))
    for ma in out:
        out[ma].sort(key=lambda mot: mot[0])
    return out


def ma_do(scene: Scene, model) -> list:
    if "do" in scene.truong:
        return [ma.strip() for ma in scene.truong["do"][0].split(",") if ma.strip()]
    return [dl["ma"] for dl in model.khai_bao["daiLuongDo"][:2]]


def _kiem_thi_nghiem(scene: Scene, thu_muc: Path) -> None:
    mau = scene.truong["mau"][0]
    if mau == thu_vien.NEW_MODEL:
        raise CanhError(scene.so, "Video chỉ dùng mẫu có sẵn trong thư viện thí nghiệm, không dùng `moi`. Mẫu có: "
                        + ", ".join(thu_vien.list_models()) + ".")
    if mau not in thu_vien.list_models():
        raise CanhError(scene.so, f"không có mẫu `{mau}` trong thư viện thí nghiệm. Mẫu có: "
                        + ", ".join(thu_vien.list_models()) + ".")
    try:
        model = thu_vien.load(mau, thu_muc)
    except thu_vien.ModelError as exc:
        raise CanhError(scene.so, str(exc)) from exc
    lich = tham_so_theo_thoi_gian(scene)
    dong_theo_ma: dict = {}
    for value, no in zip(scene.truong.get("tham-so", []), scene.dong_truong.get("tham-so", [])):
        dong_theo_ma.setdefault(value.split()[1], []).append((no, float(value.split()[2])))
    for ma, cac_dong in dong_theo_ma.items():
        ts = model.tham_so(ma)
        if ts is None:
            co = ", ".join(t["ma"] for t in model.khai_bao["thamSo"])
            raise CanhError(scene.so, f"tham số `{ma}` không có trong mẫu `{mau}` (dòng {cac_dong[0][0]}). Có: {co}.")
        if ts.get("kieu") != "so":
            raise CanhError(scene.so, f"tham số `{ma}` không phải số nên chưa dùng được trong video.")
        for no, gia_tri in cac_dong:
            if not ts["min"] <= gia_tri <= ts["max"]:
                raise ParseError(no, f"`{ma}` = {gia_tri:g} ngoài khoảng cho phép {ts['min']:g}–{ts['max']:g}.")
    if len(lich) > MAX_THAM_SO:
        raise CanhError(scene.so, f"chỉ đổi tối đa {MAX_THAM_SO} tham số khác nhau trong một cảnh (đang có {len(lich)}).")
    codes = ma_do(scene, model)
    if len(codes) > MAX_DO:
        raise CanhError(scene.so, f"`do` chỉ liệt kê tối đa {MAX_DO} đại lượng (đang có {len(codes)}).")
    for ma in codes:
        if model.dai_luong(ma) is None:
            co = ", ".join(dl["ma"] for dl in model.khai_bao["daiLuongDo"])
            raise CanhError(scene.so, f"đại lượng đo `{ma}` không có trong mẫu `{mau}`. Có: {co}.")


def _kiem_hinh_anh(scene: Scene, thu_muc: Path) -> None:
    if scene.loai == "minh-hoa":
        for value, no in zip(scene.truong["hinh"], scene.dong_truong["hinh"]):
            try:
                ten, _nhan = hinh.tach_minh_hoa(value)
                hinh.doc(ten)
            except hinh.HinhError as exc:
                raise CanhError(scene.so, f"{exc} (dòng {no}).") from exc
        return
    if "hinh" in scene.truong:
        value, no = scene.truong["hinh"][0], scene.dong_truong["hinh"][0]
        try:
            hinh.doc(value)
        except hinh.HinhError as exc:
            raise CanhError(scene.so, f"{exc} (dòng {no}).") from exc
    if "anh" in scene.truong:
        value, no = scene.truong["anh"][0], scene.dong_truong["anh"][0]
        nguon_tay = scene.truong.get("nguon", [None])[0]
        try:
            anh.doc(thu_muc, value, nguon_tay)
        except anh.AnhError as exc:
            raise CanhError(scene.so, f"{exc} (dòng {no}).") from exc


def kiem(video: Video, thu_muc: Path) -> list:
    warnings: list = []
    for scene in video.canh:
        for key, values in scene.truong.items():
            gioi_han = LIMITS.get((scene.loai, key))
            if gioi_han is None:
                continue
            for value, no in zip(values, scene.dong_truong[key]):
                kiem_danh_dau(key, value, no)
                so_ky_tu = hien_thi(value)
                if so_ky_tu > gioi_han:
                    raise CanhError(scene.so, f"`{key}` dài {so_ky_tu} ký tự, tối đa {gioi_han} (dòng {no}). Rút gọn nội dung.")
        if len(scene.loi) > LOI_DAI:
            warnings.append(f"Cảnh {scene.so}: lời dài {len(scene.loi)} ký tự (quá {LOI_DAI}); nên tách thành hai cảnh.")
        if scene.loai == "thi-nghiem":
            _kiem_thi_nghiem(scene, thu_muc)
        _kiem_hinh_anh(scene, thu_muc)
    return warnings
