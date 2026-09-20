"""Bản tính lại độc lập bằng Python của 8 mô hình thí nghiệm ảo.

Dùng cho hai việc: test so với JavaScript, và bảng số liệu lí tưởng ở trang giáo viên của phiếu học tập.
Mỗi hàm nhận dict tham số (đã đủ mặc định) và trả về dict đại lượng đo.
"""

from __future__ import annotations

import math

MASK = 0xFFFFFFFF
KW = 1.0e-14
KA = {"ch3cooh": 1.75e-5}
DELTA_H = 57200.0
DELTA_S = 175.83
R = 8.314462618
R_BAR = 0.08314462618
K25 = 0.00125
T25 = 298.15
EA = {"khong": 50000.0, "co": 45000.0}
LUONG_PHAN_UNG = 0.005
XAC_SUAT = {
    "dong-xu-ngua": 1 / 2, "xuc-xac-mat-6": 1 / 6, "xuc-xac-chan": 1 / 2,
    "hai-xuc-xac-tong-7": 1 / 6, "hai-xuc-xac-tong-12": 1 / 36,
}


def _imul(x: int, y: int) -> int:
    return (x * y) & MASK


def ngau_nhien(hat_giong: int):
    """mulberry32, khớp từng bit với taoNgauNhien trong khung.js."""
    state = hat_giong & MASK

    def rng() -> float:
        nonlocal state
        state = (state + 0x6D2B79F5) & MASK
        t = _imul(state ^ (state >> 15), 1 | state)
        t = ((t + _imul(t ^ (t >> 7), 61 | t)) & MASK) ^ t
        return ((t ^ (t >> 14)) & MASK) / 4294967296

    return rng


def li_nem_xien(p: dict) -> dict:
    goc = math.radians(p["goc"])
    vy = p["van-toc-dau"] * math.sin(goc)
    vx = p["van-toc-dau"] * math.cos(goc)
    thoi_gian = (vy + math.sqrt(vy * vy + 2 * p["g"] * p["do-cao-dau"])) / p["g"]
    return {
        "thoi-gian-bay": thoi_gian,
        "tam-xa": vx * thoi_gian,
        "do-cao-cuc-dai": p["do-cao-dau"] + vy * vy / (2 * p["g"]),
    }


def li_con_lac_don(p: dict) -> dict:
    chu_ki = 2 * math.pi * math.sqrt(p["chieu-dai"] / p["g"])
    return {"chu-ki": chu_ki, "thoi-gian-10-dao-dong": 10 * chu_ki}


def li_mach_ohm(p: dict) -> dict:
    e, r1, r2 = p["suat-dien-dong"], p["dien-tro-1"], p["dien-tro-2"]
    if p["kieu-mac"] == "song-song":
        return {
            "dien-tro-tuong-duong": 1 / (1 / r1 + 1 / r2),
            "cuong-do-mach-chinh": e / r1 + e / r2,
            "cuong-do-1": e / r1, "cuong-do-2": e / r2,
            "hieu-dien-the-1": e, "hieu-dien-the-2": e,
        }
    dong = e / (r1 + r2)
    return {
        "dien-tro-tuong-duong": r1 + r2,
        "cuong-do-mach-chinh": dong,
        "cuong-do-1": dong, "cuong-do-2": dong,
        "hieu-dien-the-1": dong * r1, "hieu-dien-the-2": dong * r2,
    }


def hoa_chuan_do(p: dict) -> dict:
    tong = p["the-tich-acid"] + p["the-tich-base"]
    acid = p["nong-do-acid"] * p["the-tich-acid"] / tong
    natri = p["nong-do-base"] * p["the-tich-base"] / tong
    if p["loai-acid"] == "hcl":
        lech = acid - natri
        h = (lech + math.sqrt(lech * lech + 4 * KW)) / 2
        return {"ph": -math.log10(h)}
    ka = KA[p["loai-acid"]]

    # Khác bản JS: khử mẫu thành đa thức bậc ba theo [H+], rồi chia đôi trên log10[H+].
    def da_thuc(h: float) -> float:
        return h ** 3 + (ka + natri) * h ** 2 + (ka * natri - KW - acid * ka) * h - ka * KW

    thap, cao = -14.0, 0.0
    for _ in range(200):
        giua = (thap + cao) / 2
        if da_thuc(10 ** giua) < 0:
            thap = giua
        else:
            cao = giua
    return {"ph": -(thap + cao) / 2}


def hoa_can_bang_no2(p: dict) -> dict:
    nhiet_do = p["nhiet-do"] + 273.15
    ap_suat = p["ap-suat"]
    kp = math.exp(-(DELTA_H - nhiet_do * DELTA_S) / (R * nhiet_do))
    # Giải theo độ phân li alpha rồi suy ra phần mol: cách khác với bản JS (JS giải thẳng phần mol).
    alpha = math.sqrt(kp / (kp + 4 * ap_suat))
    phan_mol = 2 * alpha / (1 + alpha)
    return {
        "kp": kp,
        "phan-mol-no2": phan_mol,
        "do-phan-li": alpha,
        "nong-do-no2": phan_mol * ap_suat / (R_BAR * nhiet_do),
    }


def hoa_toc_do(p: dict) -> dict:
    nhiet_do = p["nhiet-do"] + 273.15
    k = K25 * math.exp(EA["khong"] / (R * T25) - EA[p["xuc-tac"]] / (R * nhiet_do))
    return {"thoi-gian": LUONG_PHAN_UNG / (k * p["nong-do"])}


def toan_ham_so(p: dict) -> dict:
    a, b, c, d, x = p["a"], p["b"], p["c"], p["d"], p["x0"]
    cuc_dai = cuc_tieu = None
    nghiem = []
    if a != 0:
        biet = (2 * b) ** 2 - 4 * (3 * a) * c
        if biet > 0:
            nghiem = sorted(((-2 * b - math.sqrt(biet)) / (6 * a), (-2 * b + math.sqrt(biet)) / (6 * a)))
    elif b != 0:
        nghiem = [-c / (2 * b)]
    for diem in nghiem:
        dao_ham_cap_hai = 6 * a * diem + 2 * b
        if dao_ham_cap_hai < 0:
            cuc_dai = diem
        elif dao_ham_cap_hai > 0:
            cuc_tieu = diem
    return {
        "gia-tri": a * x ** 3 + b * x ** 2 + c * x + d,
        "he-so-goc": 3 * a * x ** 2 + 2 * b * x + c,
        "hoanh-do-cuc-dai": cuc_dai,
        "hoanh-do-cuc-tieu": cuc_tieu,
    }


def toan_xac_suat(p: dict) -> dict:
    rng = ngau_nhien(int(p["hat-giong"]))
    phep_thu = p["phep-thu"]
    tan_so = 0
    for _ in range(int(p["so-lan"])):
        if phep_thu == "dong-xu-ngua":
            xay_ra = rng() < 0.5
        else:
            mat = math.floor(rng() * 6) + 1
            if phep_thu == "xuc-xac-mat-6":
                xay_ra = mat == 6
            elif phep_thu == "xuc-xac-chan":
                xay_ra = mat % 2 == 0
            else:
                tong = mat + math.floor(rng() * 6) + 1
                xay_ra = tong == (7 if phep_thu == "hai-xuc-xac-tong-7" else 12)
        tan_so += xay_ra
    return {"tan-so": tan_so, "tan-suat": tan_so / p["so-lan"], "xac-suat-li-thuyet": XAC_SUAT[phep_thu]}


THAM_CHIEU = {
    "li-nem-xien": li_nem_xien,
    "li-con-lac-don": li_con_lac_don,
    "li-mach-ohm": li_mach_ohm,
    "hoa-chuan-do": hoa_chuan_do,
    "hoa-can-bang-no2": hoa_can_bang_no2,
    "hoa-toc-do": hoa_toc_do,
    "toan-ham-so": toan_ham_so,
    "toan-xac-suat": toan_xac_suat,
}
