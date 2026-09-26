"""Lịch thời gian: tách câu, mốc câu, mốc hiện ý, thời lượng cảnh, dữ liệu JSON cho trang. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path

from .kiem import CanhError, ma_do, tham_so_theo_thoi_gian
from .parse import Scene

FPS = 30
DAN_DAU = 1.0
LAU_BANG = 0.5
DUOI = 0.6
TOI_THIEU = 2.5
CANH_DAI = 40.0
VIDEO_DAI = 480.0
_CAU_RE = re.compile(r"(?<=[.!?…])\s+")
_KHOA_RE = re.compile(r"[^\w\s]", re.UNICODE)


@dataclass
class GiongInfo:
    mp3: Path | None
    giay: float
    moc_cau: list
    uoc_luong: bool
    nguon: str
    moc_tu: list = field(default_factory=list)
    uoc_luong_tu: bool = True


@dataclass
class CanhLich:
    so: int
    bat_dau: float
    thoi_luong: float
    so_khung: int
    giay_giong: float
    cau: list
    moc_cau: list
    moc_cau_giong: list
    uoc_luong: bool
    moc_tu: list = field(default_factory=list)


def tach_cau(loi: str) -> list:
    return [cau.strip() for cau in _CAU_RE.split(loi.strip()) if cau.strip()]


def moc_uoc_luong(cau: list, giay: float) -> list:
    tong = sum(len(c) for c in cau) or 1
    da_qua = 0
    moc = []
    for c in cau:
        moc.append(round(da_qua / tong * giay, 3))
        da_qua += len(c)
    return moc


def _khoa(chu: str) -> str:
    return _KHOA_RE.sub("", chu).lower()


def moc_tu_uoc_luong(loi: str, moc_cau_giong: list, giay: float) -> list:
    bien = list(moc_cau_giong) + [giay]
    ket: list = []
    for i, c in enumerate(tach_cau(loi)):
        tu = c.split()
        if not tu:
            continue
        start = bien[i] if i < len(bien) else giay
        end = bien[i + 1] if i + 1 < len(bien) else giay
        dai = max(end - start, 0.0)
        tong = sum(len(t) for t in tu) or 1
        da_qua = 0
        for t in tu:
            t0 = start + da_qua / tong * dai
            da_qua += len(t)
            t1 = start + da_qua / tong * dai
            ket.append({"t": round(t0, 3), "d": round(max(t1 - t0, 0.0), 3), "chu": t})
    return ket


def thoi_luong_canh(giay_giong: float, fps: int = FPS) -> float:
    tho = max(TOI_THIEU, DAN_DAU + giay_giong + DUOI)
    return math.ceil(tho * fps - 1e-9) / fps


def moc_hien(so_muc: int, moc_cau_giong: list, giay: float) -> list:
    if so_muc <= 0:
        return []
    if len(moc_cau_giong) >= so_muc:
        return list(moc_cau_giong[:so_muc])
    return [round(k * giay / so_muc, 3) for k in range(so_muc)]


def so_muc(scene: Scene) -> int:
    t = scene.truong
    return {
        "y-tung-y": len(t.get("y", [])),
        "quy-trinh": len(t.get("buoc", [])),
        "so-sanh": len(t.get("y-trai", [])) + len(t.get("y-phai", [])),
        "do-thi": len(t.get("diem", [])),
        "cong-thuc": len(t.get("giai-thich", [])),
        "minh-hoa": len(t.get("hinh", [])),
    }.get(scene.loai, 0)


def dung_lich(cac_canh: list, cac_giong: list, fps: int = FPS, kiem_moc: bool = True) -> tuple:
    plan: list = []
    warnings: list = []
    bat_dau = 0.0
    for scene, giong in zip(cac_canh, cac_giong):
        cau = tach_cau(scene.loi)
        uoc = giong.uoc_luong or len(giong.moc_cau) != len(cau)
        moc_giong = moc_uoc_luong(cau, giong.giay) if uoc else list(giong.moc_cau)
        if uoc:
            warnings.append(f"Cảnh {scene.so}: mốc câu ước lượng theo số ký tự; hình có thể lệch tiếng vài trăm mili giây.")
        uoc_tu = giong.uoc_luong_tu or not giong.moc_tu
        tu_tho = moc_tu_uoc_luong(scene.loi, moc_giong, giong.giay) if uoc_tu else list(giong.moc_tu)
        if uoc_tu:
            warnings.append(f"Cảnh {scene.so}: mốc từng từ ước lượng theo tỉ lệ ký tự; nhấn ý và phụ đề karaoke có thể lệch tiếng vài trăm mili giây.")
        moc_tu = [{"t": round(DAN_DAU + w["t"], 3), "d": round(w.get("d", 0.0), 3), "chu": w["chu"], "khoa": _khoa(w["chu"])}
                  for w in tu_tho]
        thoi_luong = thoi_luong_canh(giong.giay, fps)
        if thoi_luong > CANH_DAI:
            warnings.append(f"Cảnh {scene.so}: dài {thoi_luong:.0f} giây (quá 40 giây); nên tách thành hai cảnh.")
        if kiem_moc and scene.loai == "thi-nghiem":
            for value, no in zip(scene.truong.get("tham-so", []), scene.dong_truong.get("tham-so", [])):
                giay = float(value.split()[0])
                if giay > thoi_luong:
                    raise CanhError(scene.so, f"mốc `tham-so` {giay:g} giây (dòng {no}) vượt thời lượng cảnh {thoi_luong:.1f} giây.")
        plan.append(CanhLich(
            so=scene.so, bat_dau=bat_dau, thoi_luong=thoi_luong, so_khung=round(thoi_luong * fps),
            giay_giong=giong.giay, cau=cau, moc_cau=[round(DAN_DAU + m, 3) for m in moc_giong],
            moc_cau_giong=moc_giong, uoc_luong=uoc, moc_tu=moc_tu,
        ))
        bat_dau += thoi_luong
    if bat_dau > VIDEO_DAI:
        warnings.append(f"Video dài {bat_dau / 60:.1f} phút (quá 8 phút); nên tách thành nhiều video.")
    return plan, warnings


def du_lieu_canh(scene: Scene, cl: CanhLich, model=None, tai_nguyen: dict | None = None) -> dict:
    tai_nguyen = tai_nguyen or {}
    meta = tai_nguyen.get("meta", {})
    du = {
        "so": scene.so,
        "loai": scene.loai,
        "thoiLuong": cl.thoi_luong,
        "danDau": DAN_DAU,
        "giayLauBang": LAU_BANG,
        "truong": scene.truong,
        "moc": [round(DAN_DAU + m, 3) for m in moc_hien(so_muc(scene), cl.moc_cau_giong, cl.giay_giong)],
        "tu": list(cl.moc_tu),
        "hinh": tai_nguyen.get("hinh"),
        "anh": tai_nguyen.get("anh"),
        "hinhs": tai_nguyen.get("hinhs", []),
        "co": {
            "banTay": meta.get("ban-tay", "co") == "co",
            "mayQuay": meta.get("may-quay", "co") == "co",
            "lauBang": meta.get("chuyen-canh", "lau-bang") == "lau-bang" and scene.so > 1,
        },
        "nenTruoc": None,
    }
    if scene.loai == "do-thi":
        du["diem"] = [[float(p) for p in v.split(",")] for v in scene.truong["diem"]]
    if scene.loai == "thi-nghiem":
        du["khaiBao"] = model.khai_bao
        du["thamSo"] = {ma: [[g, v] for g, v in ds] for ma, ds in tham_so_theo_thoi_gian(scene).items()}
        du["do"] = ma_do(scene, model)
    return du
