"""Kiểm giới hạn chữ và nội dung cảnh của video.md. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re
from pathlib import Path

from thi_nghiem_parts import thu_vien

from .parse import ParseError, Scene, Video

LIMITS = {
    ("tieu-de", "chu"): 90, ("tieu-de", "phu"): 90,
    ("khai-niem", "thuat-ngu"): 60, ("khai-niem", "dinh-nghia"): 220,
    ("cong-thuc", "bieu-thuc"): 90, ("cong-thuc", "giai-thich"): 80,
    ("y-tung-y", "tieu-de"): 90, ("y-tung-y", "y"): 80,
    ("quy-trinh", "tieu-de"): 90, ("quy-trinh", "buoc"): 50,
    ("so-sanh", "tieu-de"): 90, ("so-sanh", "trai"): 24, ("so-sanh", "phai"): 24,
    ("so-sanh", "y-trai"): 60, ("so-sanh", "y-phai"): 60,
    ("do-thi", "tieu-de"): 90, ("do-thi", "truc-ngang"): 40, ("do-thi", "truc-doc"): 40,
}
LOI_DAI = 700
MAX_THAM_SO = 3
MAX_DO = 3
_MARKUP_RE = re.compile(r"\*\*|~|\^")


class CanhError(Exception):
    """Nội dung cảnh sai; luôn nêu số cảnh."""

    def __init__(self, so: int, message: str) -> None:
        super().__init__(f"Cảnh {so}: {message}")
        self.so = so
        self.message = message


def hien_thi(chu: str) -> int:
    return len(_MARKUP_RE.sub("", chu))


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


def kiem(video: Video, thu_muc: Path) -> list:
    warnings: list = []
    for scene in video.canh:
        for key, values in scene.truong.items():
            gioi_han = LIMITS.get((scene.loai, key))
            if gioi_han is None:
                continue
            for value, no in zip(values, scene.dong_truong[key]):
                so_ky_tu = hien_thi(value)
                if so_ky_tu > gioi_han:
                    raise CanhError(scene.so, f"`{key}` dài {so_ky_tu} ký tự, tối đa {gioi_han} (dòng {no}). Rút gọn nội dung.")
        if len(scene.loi) > LOI_DAI:
            warnings.append(f"Cảnh {scene.so}: lời dài {len(scene.loi)} ký tự (quá {LOI_DAI}); nên tách thành hai cảnh.")
        if scene.loai == "thi-nghiem":
            _kiem_thi_nghiem(scene, thu_muc)
    return warnings
