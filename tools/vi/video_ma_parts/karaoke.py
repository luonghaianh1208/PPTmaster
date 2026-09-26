"""Phụ đề karaoke: file .ass với \\kf theo mốc từng từ, dùng font Itim. Chỉ dùng thư viện chuẩn."""

from __future__ import annotations

import re

from .lich import DAN_DAU

GIOI_HAN_KY_TU = 42
_MARKUP_RE = re.compile(r"\*\*|~|\^|==|\(\(|\)\)|__|\{\{|\}\}")
_ESCAPE = (("\\", "\\\\"), ("{", "\\{"), ("}", "\\}"))

_HEADER = """[Script Info]
Title: Phụ đề karaoke
ScriptType: v4.00+
PlayResX: {rong}
PlayResY: {cao}
WrapStyle: 0
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Itim,Itim,40,&H0000D7FF,&H00FFFFFF,&H00000000,&H00000000,0,0,0,0,100,100,1.25,0,1,3,0,2,10,10,55,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""


def _thoat(chu: str) -> str:
    for cu, moi in _ESCAPE:
        chu = chu.replace(cu, moi)
    return chu


def _thoi_gian(giay: float) -> str:
    giay = max(giay, 0.0)
    tong_cs = int(round(giay * 100))
    gio, du = divmod(tong_cs, 360000)
    phut, du = divmod(du, 6000)
    s, cs = divmod(du, 100)
    return f"{gio}:{phut:02d}:{s:02d}.{cs:02d}"


def _boc_dong(chu: list, gioi_han: int = GIOI_HAN_KY_TU) -> list:
    """Chia chỉ số từ (0..n-1) thành các dòng, mỗi dòng tối đa `gioi_han` ký tự hiển thị, không cắt giữa từ."""
    dong: list = []
    hien: list = []
    dai = 0
    for i, chu_i in enumerate(chu):
        them = len(chu_i) + (1 if hien else 0)
        if hien and dai + them > gioi_han:
            dong.append(hien)
            hien, dai = [], 0
            them = len(chu_i)
        hien.append(i)
        dai += them
    if hien:
        dong.append(hien)
    return dong or [[]]


def _nhom_cau(cl) -> list:
    """Gom `cl.moc_tu` theo từng câu bằng mốc `cl.moc_cau`; trả `[(cau_text, start, end, [tu...])]` (thời gian trong cảnh)."""
    if not cl.cau:
        return []
    bien = list(cl.moc_cau[1:]) + [DAN_DAU + cl.giay_giong]
    nhom = [[] for _ in cl.cau]
    idx = 0
    for w in cl.moc_tu:
        while idx < len(bien) - 1 and w["t"] >= bien[idx] - 1e-6:
            idx += 1
        nhom[idx].append(w)
    return [(cl.cau[k], cl.moc_cau[k], bien[k], nhom[k]) for k in range(len(cl.cau))]


def _kf_cs(neo: float, moc: list) -> list:
    """`moc`: n+1 mốc giây (đầu = neo), trả n khoảng centi-giây không âm, không giảm, cộng lại đúng tổng."""
    cs = [max(round((m - neo) * 100), 0) for m in moc]
    for i in range(1, len(cs)):
        cs[i] = max(cs[i], cs[i - 1])
    return [cs[i + 1] - cs[i] for i in range(len(cs) - 1)]


def _dialogue(bat_dau_canh: float, start: float, end: float, style: str = "Itim") -> str:
    return (f"Dialogue: 0,{_thoi_gian(bat_dau_canh + start)},{_thoi_gian(bat_dau_canh + end)},"
            f"{style},,0,0,0,,")


def _dialogues_cau(bat_dau_canh: float, start: float, end: float, tu: list, cau_text: str) -> list:
    if not tu:
        text = _thoat(_MARKUP_RE.sub("", cau_text))
        return [_dialogue(bat_dau_canh, start, end) + text]
    texts = [_thoat(_MARKUP_RE.sub("", w["chu"])) for w in tu]
    hien = [_MARKUP_RE.sub("", w["chu"]) for w in tu]
    moc = [w["t"] for w in tu] + [end]
    dong = _boc_dong(hien)
    ket: list = []
    for i in range(0, len(dong), 2):
        nhom_dong = dong[i:i + 2]
        chi_so = [j for dong_k in nhom_dong for j in dong_k]
        g_start, g_end = moc[chi_so[0]], moc[chi_so[-1] + 1]
        kf = _kf_cs(g_start, [moc[j] for j in chi_so] + [g_end])
        parts = []
        pos = 0
        for li, dong_k in enumerate(nhom_dong):
            for j2, idx in enumerate(dong_k):
                dai = kf[pos]
                pos += 1
                cuoi_dong = j2 == len(dong_k) - 1
                if cuoi_dong:
                    hau_to = "\\N" if li < len(nhom_dong) - 1 else ""
                else:
                    hau_to = " "
                parts.append(f"{{\\kf{dai}}}{texts[idx]}{hau_to}")
        ket.append(_dialogue(bat_dau_canh, g_start, g_end) + "".join(parts))
    return ket


def tao_ass(cac_lich: list, rong: int = 1280, cao: int = 720) -> str:
    dialogues: list = []
    for cl in cac_lich:
        for cau_text, start, end, tu in _nhom_cau(cl):
            dialogues.extend(_dialogues_cau(cl.bat_dau, start, end, tu, cau_text))
    return _HEADER.format(rong=rong, cao=cao) + "\n".join(dialogues) + ("\n" if dialogues else "")
