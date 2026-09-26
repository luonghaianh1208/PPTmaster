"""Phụ đề karaoke: file .ass với \\kf theo mốc từng từ, dùng font Itim. Chỉ dùng thư viện chuẩn.

Chữ hiển thị luôn là token gốc của kịch bản (đã bỏ đánh dấu nhấn nhưng giữ nguyên dấu câu, chữ hoa/thường);
mốc thời gian của từng token lấy từ sự kiện giọng máy (`tu`) bằng đúng phép khớp chữ mà `giong.py` dùng cho
mốc đầu câu (`giong.can_chinh_tu`), vì một token kịch bản có thể trải trên nhiều sự kiện (số đọc từng chữ
số) hoặc nhiều token gộp vào một sự kiện. Không khớp chắc chắn được thì chia đều theo tỉ lệ ký tự, như
đường ước lượng của `lich.moc_tu_uoc_luong`.
"""

from __future__ import annotations

import re

from . import giong
from .lich import doan_loi

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


def _boc_dong(tokens: list, gioi_han: int = GIOI_HAN_KY_TU) -> list:
    """Chia chỉ số token (0..n-1) thành các dòng, mỗi dòng tối đa `gioi_han` ký tự hiển thị, không cắt giữa từ."""
    dong: list = []
    hien: list = []
    dai = 0
    for i, tok in enumerate(tokens):
        them = len(tok) + (1 if hien else 0)
        if hien and dai + them > gioi_han:
            dong.append(hien)
            hien, dai = [], 0
            them = len(tok)
        hien.append(i)
        dai += them
    if hien:
        dong.append(hien)
    return dong or [[]]


def _do_dai_dong(chi_so: list, tokens: list) -> int:
    return sum(len(tokens[i]) for i in chi_so) + len(chi_so) - 1


def _can_bang_hai_dong(nhom_dong: list, tokens: list, gioi_han: int = GIOI_HAN_KY_TU) -> list:
    """Cân bằng lại ranh giới giữa 2 dòng của cùng một Dialogue (độ dài ~ số ký tự) để dòng 2 không mồ côi:
    khi gộp lại có ít nhất 4 từ, dòng 2 phải có ít nhất 2 từ. Không đổi tập token, chỉ đổi điểm cắt."""
    if len(nhom_dong) != 2:
        return nhom_dong
    chi_so = nhom_dong[0] + nhom_dong[1]
    if len(chi_so) < 4:
        return nhom_dong
    tot_nhat = None
    for k in range(2, len(chi_so) - 1):
        d1, d2 = chi_so[:k], chi_so[k:]
        l1, l2 = _do_dai_dong(d1, tokens), _do_dai_dong(d2, tokens)
        if l1 > gioi_han or l2 > gioi_han:
            continue
        lech = abs(l1 - l2)
        if tot_nhat is None or lech < tot_nhat[0]:
            tot_nhat = (lech, d1, d2)
    return nhom_dong if tot_nhat is None else [tot_nhat[1], tot_nhat[2]]


def _chia_ti_le(tokens: list, start: float, end: float) -> list:
    """Chia đều theo tỉ lệ ký tự trong `[start, end)`, như `lich.moc_tu_uoc_luong` — dùng khi không có sự
    kiện giọng máy, hoặc khi khớp chữ với sự kiện không đủ tin cậy."""
    tong = sum(len(t) for t in tokens) or 1
    da_qua = 0
    ket = []
    for t in tokens:
        ket.append(start + da_qua / tong * (end - start))
        da_qua += len(t)
    return ket


def _noi_suy_theo_su_kien(idx_su_kien: list, tokens: list, tu: list, start: float, end: float) -> list:
    """Mốc bắt đầu từng token: token khớp trực tiếp với sự kiện thì lấy đúng giờ sự kiện đó; token bị gộp
    vào sự kiện của token liền trước (`idx_su_kien[k] is None`) thì chia đều theo tỉ lệ ký tự giữa hai mốc
    đã biết gần nhất (hoặc `start`/`end` ở hai đầu)."""
    n = len(tokens)
    biet = [tu[idx]["t"] if idx is not None else None for idx in idx_su_kien]
    ket: list = [None] * n
    k = 0
    while k < n:
        if biet[k] is not None:
            ket[k] = biet[k]
            k += 1
            continue
        m = k
        while m < n and biet[m] is None:
            m += 1
        truoc_t = biet[k - 1] if k > 0 else start
        sau_t = biet[m] if m < n else end
        tong = sum(len(tokens[x]) for x in range(k, m)) or 1
        da_qua = 0
        for x in range(k, m):
            ket[x] = truoc_t + da_qua / tong * (sau_t - truoc_t)
            da_qua += len(tokens[x])
        k = m
    return ket


def _mocs_tu_kich_ban(tokens: list, tu: list, start: float, end: float) -> list:
    if not tokens:
        return []
    if not tu:
        return _chia_ti_le(tokens, start, end)
    chuan_tu = [giong._chuan_hoa(t) for t in tokens]
    chuan_su_kien = [giong._chuan_hoa(w["chu"]) for w in tu]
    idx_su_kien, _i, j = giong.can_chinh_tu(chuan_tu, chuan_su_kien)
    khop = idx_su_kien[0] is not None and abs(len(tu) - j) <= max(3, round(len(tu) * 0.1))
    if not khop:
        return _chia_ti_le(tokens, start, end)
    return _noi_suy_theo_su_kien(idx_su_kien, tokens, tu, start, end)


def _nhom_cau(cl) -> list:
    """Gom mốc từ theo từng câu của từng đoạn lời (`lich.doan_loi`: lời, và lời giải của cảnh câu hỏi);
    trả `[(cau_text, start, end, [tu...])]` (thời gian trong cảnh)."""
    ket: list = []
    for cau, moc_cau, het, moc_tu in doan_loi(cl):
        if not cau:
            continue
        bien = list(moc_cau[1:]) + [het]
        nhom = [[] for _ in cau]
        idx = 0
        for w in moc_tu:
            while idx < len(bien) - 1 and w["t"] >= bien[idx] - 1e-6:
                idx += 1
            nhom[idx].append(w)
        ket.extend((cau[k], moc_cau[k], bien[k], nhom[k]) for k in range(len(cau)))
    return ket


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
    tokens = _MARKUP_RE.sub("", cau_text).split()
    if not tokens:
        return []
    texts = [_thoat(t) for t in tokens]
    moc_bat_dau = _mocs_tu_kich_ban(tokens, tu, start, end)
    moc = moc_bat_dau + [end]
    dong = _boc_dong(tokens)
    ket: list = []
    for i in range(0, len(dong), 2):
        nhom_dong = _can_bang_hai_dong(dong[i:i + 2], tokens)
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
