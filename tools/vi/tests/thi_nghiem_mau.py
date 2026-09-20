"""Dữ liệu mẫu dùng chung cho các test thí nghiệm ảo."""

import random

PENDULUM = """---
tieu-de: Chu kì con lắc đơn
mon: Vật lí
lop: 11
mau: li-con-lac-don
nguoi-thao-tac: nhom
sai-so: bat
---

## Tham số
chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0
g: co-dinh 9.8

## Dự đoán
cau: Khi tăng chiều dài dây gấp 4 lần, chu kì thay đổi thế nào?
A: Tăng 4 lần
B: Tăng 2 lần
C: Không đổi
dap-an: B

## Quan sát
so-lan-do: 5
do: chieu-dai, chu-ki
do-thi: chu-ki^2 theo chieu-dai

## Giải thích
cau: Từ đồ thị, T^2^ và l liên hệ thế nào?
goi-y-dap-an: T^2^ tỉ lệ thuận với l; T = 2π√(l/g).

## Kết luận
Chu kì con lắc đơn chỉ phụ thuộc chiều dài dây và g.
"""

HOOKE_JSON = {
    "ma": "moi", "ten": "Định luật Hooke", "mon": "Vật lí", "hoatHinh": "khong",
    "thamSo": [
        {"ma": "do-cung", "ten": "Độ cứng k", "kieu": "so", "donVi": "N/m", "min": 10, "max": 100, "buoc": 10, "macDinh": 50},
        {"ma": "do-gian", "ten": "Độ giãn x", "kieu": "so", "donVi": "m", "min": 0, "max": 0.2, "buoc": 0.01, "macDinh": 0.1},
    ],
    "daiLuongDo": [{"ma": "luc", "ten": "Lực đàn hồi F", "donVi": "N", "saiSo": 0.05, "chuSo": 2}],
    "congThuc": {"bieuThuc": "F = k·x", "dieuKien": "Lò xo còn trong giới hạn đàn hồi.", "nguon": ""},
    "bangKiem": [
        {"vao": {}, "ra": {"luc": 5.0}, "saiSo": 0.001},
        {"vao": {"do-gian": 0}, "ra": {"luc": 0.0}, "saiSo": 0.001},
        {"vao": {"do-gian": 0.2}, "ra": {"luc": 10.0}, "saiSo": 0.001},
        {"vao": {"do-cung": 100}, "ra": {"luc": 10.0}, "saiSo": 0.001},
        {"vao": {"do-cung": 10, "do-gian": 0.05}, "ra": {"luc": 0.5}, "saiSo": 0.001},
    ],
}
HOOKE_JS = """(function (root) {
  function tinh(p) { return { 'luc': p['do-cung'] * p['do-gian'] }; }
  function ve(ctx, p, t, kt, d) { ctx.fillRect(20, 20, 40 + 600 * p['do-gian'], 20); }
  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
"""
HOOKE_MD = """---
tieu-de: Định luật Hooke
mon: Vật lí
lop: 10
mau: moi
---

## Tham số
do-gian: 0..0.2 buoc 0.05 mac-dinh 0.1

## Dự đoán
cau: Độ giãn tăng gấp đôi thì lực đàn hồi thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: do-gian, luc
do-thi: luc theo do-gian

## Giải thích
cau: Lực đàn hồi và độ giãn liên hệ thế nào?
goi-y-dap-an: F tỉ lệ thuận với x.

## Kết luận
Trong giới hạn đàn hồi, lực đàn hồi tỉ lệ thuận với độ giãn.
"""


def grid(model, count: int, rng: random.Random) -> list[dict]:
    points = []
    for _ in range(count):
        point = {}
        for ts in model.khai_bao["thamSo"]:
            if ts["kieu"] == "chon":
                point[ts["ma"]] = rng.choice(ts["luaChon"])["ma"]
            else:
                steps = round((ts["max"] - ts["min"]) / ts["buoc"])
                point[ts["ma"]] = round(ts["min"] + rng.randint(0, steps) * ts["buoc"], 6)
        points.append(point)
    return points
