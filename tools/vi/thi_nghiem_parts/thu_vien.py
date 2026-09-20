"""Nạp và kiểm mô hình thí nghiệm ảo: mẫu trong thư viện hoặc mô hình mới trong thư mục thí nghiệm.

Mỗi mô hình gồm hai file: <mã>.json (khai báo) và <mã>.js (hàm tinh, ve). Khuôn ở docs/vi/tro-ly/mo-hinh-thi-nghiem.md.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

LIBRARY_DIR = Path(__file__).resolve().parent / "mo_hinh"
NEW_MODEL = "moi"
NEW_JSON = "mo-hinh.json"
NEW_JS = "mo-hinh.js"
MIN_CHECK_ROWS = 5
ANIMATIONS = ("khong", "mot-lan", "lap")
BANNED_JS = (
    "http://", "https://", "fetch(", "XMLHttpRequest", "WebSocket", "import(", "import ", "require(",
    "eval(", "Function(", "process.", "document.", "window.", "localStorage", "</script",
)


class ModelError(Exception):
    """Mô hình không có, hoặc không đúng khuôn."""


@dataclass
class Model:
    ma: str
    khai_bao: dict
    js: str
    json_path: Path
    js_path: Path
    moi: bool

    def tham_so(self, ma: str) -> dict | None:
        return next((ts for ts in self.khai_bao["thamSo"] if ts["ma"] == ma), None)

    def dai_luong(self, ma: str) -> dict | None:
        return next((dl for dl in self.khai_bao["daiLuongDo"] if dl["ma"] == ma), None)

    def mac_dinh(self) -> dict:
        return {ts["ma"]: ts["macDinh"] for ts in self.khai_bao["thamSo"]}


def list_models() -> list[str]:
    return sorted(path.stem for path in LIBRARY_DIR.glob("*.json"))


def _number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _text(value) -> bool:
    return isinstance(value, str) and value.strip() != ""


def check_declaration(kb) -> list[str]:
    """Trả về danh sách lỗi của file khai báo; rỗng là đạt."""
    if not isinstance(kb, dict):
        return ["file khai báo phải là một đối tượng JSON"]
    errors: list[str] = []
    for key in ("ten", "mon"):
        if not _text(kb.get(key)):
            errors.append(f"thiếu `{key}`")
    if kb.get("hoatHinh") not in ANIMATIONS:
        errors.append("`hoatHinh` phải là một trong: " + ", ".join(ANIMATIONS))

    params = kb.get("thamSo")
    param_codes: dict[str, dict] = {}
    if not isinstance(params, list) or not params:
        errors.append("`thamSo` phải có ít nhất một tham số")
        params = []
    for index, ts in enumerate(params, 1):
        where = f"thamSo[{index}]"
        if not isinstance(ts, dict) or not _text(ts.get("ma")) or not _text(ts.get("ten")):
            errors.append(f"{where}: thiếu `ma` hoặc `ten`")
            continue
        param_codes[ts["ma"]] = ts
        if ts.get("kieu") == "so":
            if not all(_number(ts.get(key)) for key in ("min", "max", "buoc", "macDinh")):
                errors.append(f"{where}: tham số số cần `min`, `max`, `buoc`, `macDinh` là số")
            elif not (ts["min"] < ts["max"] and ts["buoc"] > 0 and ts["min"] <= ts["macDinh"] <= ts["max"]):
                errors.append(f"{where}: cần min < max, buoc > 0 và macDinh nằm trong khoảng")
            if not isinstance(ts.get("donVi"), str):
                errors.append(f"{where}: thiếu `donVi` (để chuỗi rỗng nếu không có đơn vị)")
        elif ts.get("kieu") == "chon":
            options = ts.get("luaChon")
            codes = [o.get("ma") for o in options if isinstance(o, dict)] if isinstance(options, list) else []
            if len(codes) < 2 or not all(_text(o.get("ma")) and _text(o.get("ten")) for o in options):
                errors.append(f"{where}: tham số lựa chọn cần ít nhất hai mục có `ma` và `ten`")
            elif ts.get("macDinh") not in codes:
                errors.append(f"{where}: `macDinh` phải là một trong các lựa chọn")
        else:
            errors.append(f"{where}: `kieu` phải là `so` hoặc `chon`")

    measures = kb.get("daiLuongDo")
    measure_codes: set[str] = set()
    if not isinstance(measures, list) or not measures:
        errors.append("`daiLuongDo` phải có ít nhất một đại lượng đo")
        measures = []
    for index, dl in enumerate(measures, 1):
        where = f"daiLuongDo[{index}]"
        if not isinstance(dl, dict) or not _text(dl.get("ma")) or not _text(dl.get("ten")):
            errors.append(f"{where}: thiếu `ma` hoặc `ten`")
            continue
        measure_codes.add(dl["ma"])
        if not isinstance(dl.get("donVi"), str):
            errors.append(f"{where}: thiếu `donVi`")
        if not _number(dl.get("saiSo")) or dl["saiSo"] < 0:
            errors.append(f"{where}: `saiSo` phải là số không âm")
        if not isinstance(dl.get("chuSo"), int) or isinstance(dl.get("chuSo"), bool) or not 0 <= dl["chuSo"] <= 8:
            errors.append(f"{where}: `chuSo` phải là số nguyên từ 0 đến 8")
    if set(param_codes) & measure_codes:
        errors.append("mã tham số và mã đại lượng đo không được trùng nhau")

    formula = kb.get("congThuc")
    if not isinstance(formula, dict) or not _text(formula.get("bieuThuc")):
        errors.append("thiếu `congThuc.bieuThuc`")
    if not isinstance(formula, dict) or not _text(formula.get("dieuKien")):
        errors.append("thiếu `congThuc.dieuKien` (điều kiện áp dụng của công thức)")

    rows = kb.get("bangKiem")
    if not isinstance(rows, list) or len(rows) < MIN_CHECK_ROWS:
        errors.append(f"`bangKiem` phải có ít nhất {MIN_CHECK_ROWS} dòng")
        rows = rows if isinstance(rows, list) else []
    for index, row in enumerate(rows, 1):
        where = f"bangKiem[{index}]"
        if not isinstance(row, dict) or not isinstance(row.get("vao"), dict) or not isinstance(row.get("ra"), dict):
            errors.append(f"{where}: cần `vao` và `ra` là đối tượng")
            continue
        if not row["ra"]:
            errors.append(f"{where}: `ra` không được rỗng")
        if not _number(row.get("saiSo")) or row["saiSo"] < 0:
            errors.append(f"{where}: `saiSo` phải là số không âm")
        for code in row["vao"]:
            if code not in param_codes:
                errors.append(f"{where}: `vao` có tham số lạ `{code}`")
        for code, value in row["ra"].items():
            if code not in measure_codes:
                errors.append(f"{where}: `ra` có đại lượng lạ `{code}`")
            elif value is not None and not _number(value):
                errors.append(f"{where}: `ra.{code}` phải là số hoặc null")
    return errors


def check_js(js: str, animation: str) -> list[str]:
    errors = [f"file mô hình không được chứa `{token.strip()}`" for token in BANNED_JS if token in js]
    for name in ("THI_NGHIEM_MO_HINH", "tinh", "ve"):
        if name not in js:
            errors.append(f"file mô hình thiếu `{name}`")
    if animation == "mot-lan" and "thoiLuong" not in js:
        errors.append("mô hình `hoatHinh: mot-lan` phải có hàm `thoiLuong`")
    return errors


def load(ma: str, folder: Path) -> Model:
    """ma là mã mẫu trong thư viện, hoặc `moi` để lấy mo-hinh.json và mo-hinh.js trong folder."""
    moi = ma == NEW_MODEL
    if moi:
        json_path, js_path = folder / NEW_JSON, folder / NEW_JS
        missing = [path.name for path in (json_path, js_path) if not path.is_file()]
        if missing:
            raise ModelError("`mau: moi` cần có trong thư mục thí nghiệm: " + ", ".join(missing))
    else:
        json_path, js_path = LIBRARY_DIR / f"{ma}.json", LIBRARY_DIR / f"{ma}.js"
        if not json_path.is_file() or not js_path.is_file():
            raise ModelError(f"Không có mẫu `{ma}`. Các mẫu có sẵn: " + ", ".join(list_models()) + "; hoặc `moi`.")
    try:
        khai_bao = json.loads(json_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ModelError(f"Không đọc được {json_path.name}: {exc}") from exc
    try:
        js = js_path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise ModelError(f"Không đọc được {js_path.name}: {exc}") from exc
    errors = check_declaration(khai_bao)
    if errors:
        raise ModelError(f"{json_path.name}: " + "; ".join(errors))
    errors = check_js(js, khai_bao["hoatHinh"])
    if errors:
        raise ModelError(f"{js_path.name}: " + "; ".join(errors))
    return Model(ma=ma, khai_bao=khai_bao, js=js, json_path=json_path, js_path=js_path, moi=moi)
