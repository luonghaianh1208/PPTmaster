"""Đọc file nguồn thi-nghiem.md thành cấu trúc dữ liệu và kiểm nó với khai báo của mô hình.

Ngữ pháp nằm trong docs/vi/tro-ly/thi-nghiem-ao.md. Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

META_REQUIRED = ("tieu-de", "mon", "lop", "mau")
META_OPTIONAL = ("nguoi-thao-tac", "sai-so")
OPERATORS = ("giao-vien", "nhom")
SECTIONS = ("## Tham số", "## Dự đoán", "## Quan sát", "## Giải thích", "## Kết luận")
OPTION_KEYS = ("A", "B", "C", "D")
MIN_RUNS, MAX_RUNS = 3, 20
MAX_COLUMNS = 6
TRANSFORMS = (
    (re.compile(r"^([a-z0-9-]+)\^2$"), "binh-phuong"),
    (re.compile(r"^1/([a-z0-9-]+)$"), "nghich-dao"),
    (re.compile(r"^ln\(([a-z0-9-]+)\)$"), "ln"),
    (re.compile(r"^sqrt\(([a-z0-9-]+)\)$"), "can"),
    (re.compile(r"^([a-z0-9-]+)$"), "khong"),
)
_NUMBER = r"-?[0-9]+(?:[.,][0-9]+)?"
_SLIDER_RE = re.compile(rf"^({_NUMBER})\.\.({_NUMBER})\s+buoc\s+({_NUMBER})\s+mac-dinh\s+({_NUMBER})$")
_KEY_RE = re.compile(r"^([A-Za-z0-9-]+):\s*(.*)$")
_URL_RE = re.compile(r"https?://", re.IGNORECASE)


class ParseError(Exception):
    """Lỗi trong thi-nghiem.md; luôn kèm số dòng để AI sửa được đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Axis:
    ma: str
    phep: str


@dataclass
class Experiment:
    meta: dict[str, str]
    tham_so: dict[str, dict]
    du_doan_cau: str
    lua_chon: list[tuple[str, str]]
    dap_an: str | None
    so_lan_do: int
    cot: list[str]
    do_thi: tuple[Axis, Axis] | None  # (trục tung, trục hoành)
    giai_thich_cau: str
    goi_y: str
    ket_luan: str
    warnings: list[str] = field(default_factory=list)

    def sliders(self) -> list[str]:
        return [ma for ma, ch in self.tham_so.items() if ch["kieu"] != "co-dinh"]

    def config(self) -> dict:
        """Cấu hình nhúng vào file HTML; khung.js đọc đúng các khoá này."""
        do_thi = None
        if self.do_thi is not None:
            tung, hoanh = self.do_thi
            do_thi = {"tung": {"ma": tung.ma, "phep": tung.phep}, "hoanh": {"ma": hoanh.ma, "phep": hoanh.phep}}
        return {
            "tieuDe": self.meta["tieu-de"], "mon": self.meta["mon"], "lop": self.meta["lop"],
            "nguoiThaoTac": self.meta["nguoi-thao-tac"], "saiSo": self.meta["sai-so"] == "bat",
            "thamSo": self.tham_so,
            "duDoan": {"cau": self.du_doan_cau,
                       "luaChon": [{"ma": ma, "noiDung": noi_dung} for ma, noi_dung in self.lua_chon],
                       "dapAn": self.dap_an},
            "quanSat": {"soLanDo": self.so_lan_do, "cot": self.cot, "doThi": do_thi},
            "giaiThich": {"cau": self.giai_thich_cau, "goiY": self.goi_y},
            "ketLuan": self.ket_luan,
        }


def _number(text: str) -> float:
    return float(text.replace(",", "."))


def _lines(text: str) -> list[tuple[int, str]]:
    lines = [(index, line.rstrip()) for index, line in enumerate(text.lstrip("﻿").splitlines(), 1)]
    for line_no, line in lines:
        if _URL_RE.search(line):
            raise ParseError(line_no, "không chèn địa chỉ web vào thí nghiệm; file thí nghiệm chạy không cần mạng")
    return lines


def _split(text: str) -> tuple[dict[str, tuple[int, str]], dict[str, list[tuple[int, str]]], dict[str, int]]:
    lines = _lines(text)
    body = [(n, line) for n, line in lines if line.strip()]
    if not body or body[0][1].strip() != "---":
        raise ParseError(body[0][0] if body else 1, "file phải mở đầu bằng khối thông tin giữa hai dòng `---`")
    meta: dict[str, tuple[int, str]] = {}
    position = lines.index(body[0]) + 1
    closed = False
    while position < len(lines):
        line_no, line = lines[position]
        position += 1
        if line.strip() == "---":
            closed = True
            break
        if not line.strip():
            continue
        match = _KEY_RE.match(line.strip())
        if match is None:
            raise ParseError(line_no, "dòng trong khối thông tin phải có dạng `khoá: giá trị`")
        key, value = match.group(1), match.group(2).strip()
        if key not in META_REQUIRED + META_OPTIONAL:
            raise ParseError(line_no, f"khoá lạ `{key}`; chỉ dùng: " + ", ".join(META_REQUIRED + META_OPTIONAL))
        if key in meta:
            raise ParseError(line_no, f"khoá `{key}` bị lặp")
        meta[key] = (line_no, value)
    if not closed:
        raise ParseError(lines[-1][0], "khối thông tin chưa đóng bằng dòng `---`")

    sections: dict[str, list[tuple[int, str]]] = {}
    heading_lines: dict[str, int] = {}
    current: str | None = None
    for line_no, line in lines[position:]:
        if line.startswith("## "):
            heading = line.strip()
            if heading not in SECTIONS:
                raise ParseError(line_no, f"mục lạ `{heading}`; chỉ dùng: " + ", ".join(SECTIONS))
            if heading in sections:
                raise ParseError(line_no, f"mục `{heading}` bị lặp")
            current = heading
            sections[current] = []
            heading_lines[current] = line_no
        elif line.strip():
            if current is None:
                raise ParseError(line_no, "nội dung phải nằm dưới một mục `## ...`")
            sections[current].append((line_no, line.strip()))
    last = lines[-1][0]
    for heading in SECTIONS:
        if heading not in sections:
            raise ParseError(last, f"thiếu mục `{heading}`")
    return meta, sections, heading_lines


def read_meta(text: str) -> dict[str, str]:
    """Đọc riêng khối thông tin, để biết `mau` trước khi nạp mô hình."""
    meta, _, _ = _split(text)
    result: dict[str, str] = {}
    for key in META_REQUIRED:
        if key not in meta or not meta[key][1]:
            raise ParseError(1, f"khối thông tin thiếu `{key}`")
        result[key] = meta[key][1]
    result["nguoi-thao-tac"] = meta.get("nguoi-thao-tac", (0, "giao-vien"))[1] or "giao-vien"
    result["sai-so"] = meta.get("sai-so", (0, "tat"))[1] or "tat"
    if result["nguoi-thao-tac"] not in OPERATORS:
        raise ParseError(meta["nguoi-thao-tac"][0], "`nguoi-thao-tac` chỉ nhận giao-vien hoặc nhom")
    if result["sai-so"] not in ("bat", "tat"):
        raise ParseError(meta["sai-so"][0], "`sai-so` chỉ nhận bat hoặc tat")
    return result


def _pairs(lines: list[tuple[int, str]], allowed: tuple[str, ...], heading: str) -> dict[str, tuple[int, str]]:
    result: dict[str, tuple[int, str]] = {}
    for line_no, line in lines:
        match = _KEY_RE.match(line)
        if match is None or match.group(1) not in allowed:
            raise ParseError(line_no, f"mục `{heading}` chỉ nhận các dòng: " + ", ".join(f"`{key}:`" for key in allowed))
        if match.group(1) in result:
            raise ParseError(line_no, f"dòng `{match.group(1)}:` bị lặp")
        result[match.group(1)] = (line_no, match.group(2).strip())
    return result


def _required(pairs: dict[str, tuple[int, str]], key: str, heading_line: int, heading: str) -> tuple[int, str]:
    if key not in pairs or not pairs[key][1]:
        raise ParseError(pairs.get(key, (heading_line, ""))[0], f"mục `{heading}` thiếu dòng `{key}:`")
    return pairs[key]


def _parameters(lines: list[tuple[int, str]], heading_line: int, khai_bao: dict) -> dict[str, dict]:
    declared = {ts["ma"]: ts for ts in khai_bao["thamSo"]}
    chosen: dict[str, dict] = {}
    for line_no, line in lines:
        match = _KEY_RE.match(line)
        if match is None:
            raise ParseError(line_no, "dòng tham số phải có dạng `<mã>: ...`")
        ma, value = match.group(1), match.group(2).strip()
        if ma not in declared:
            raise ParseError(line_no, f"mẫu không có tham số `{ma}`; các tham số của mẫu: " + ", ".join(declared))
        if ma in chosen:
            raise ParseError(line_no, f"tham số `{ma}` bị lặp")
        ts = declared[ma]
        codes = [o["ma"] for o in ts.get("luaChon", [])]
        if value.startswith("co-dinh"):
            raw = value[len("co-dinh"):].strip()
            if ts["kieu"] == "chon":
                if raw not in codes:
                    raise ParseError(line_no, f"`{ma}` chỉ nhận: " + ", ".join(codes))
                chosen[ma] = {"kieu": "co-dinh", "giaTri": raw}
            else:
                if not re.fullmatch(_NUMBER, raw):
                    raise ParseError(line_no, f"`{ma}: co-dinh` cần một số, ví dụ `co-dinh {ts['macDinh']}`")
                number = _number(raw)
                if not ts["min"] <= number <= ts["max"]:
                    raise ParseError(line_no, f"`{ma}` phải nằm trong khoảng {ts['min']}..{ts['max']} {ts['donVi']} "
                                              "(ngoài khoảng này công thức của mẫu không còn đúng)")
                chosen[ma] = {"kieu": "co-dinh", "giaTri": number}
        elif value.startswith("chon"):
            if ts["kieu"] != "chon":
                raise ParseError(line_no, f"`{ma}` là tham số số; dùng `<nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>` "
                                          "hoặc `co-dinh <giá trị>`")
            picked = [item.strip() for item in value[len("chon"):].split(",") if item.strip()]
            unknown = [item for item in picked if item not in codes]
            if len(picked) < 2 or unknown or len(set(picked)) != len(picked):
                raise ParseError(line_no, f"`{ma}: chon` cần ít nhất hai lựa chọn khác nhau trong: " + ", ".join(codes))
            chosen[ma] = {"kieu": "chon", "luaChon": picked, "macDinh": picked[0]}
        else:
            if ts["kieu"] == "chon":
                raise ParseError(line_no, f"`{ma}` là tham số lựa chọn; dùng `chon {', '.join(codes)}` hoặc `co-dinh <mã>`")
            slider = _SLIDER_RE.match(value)
            if slider is None:
                raise ParseError(line_no, f"`{ma}` cần dạng `<nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>` "
                                          "hoặc `co-dinh <giá trị>`")
            low, high, step, default = (_number(part) for part in slider.groups())
            if not ts["min"] <= low < high <= ts["max"]:
                raise ParseError(line_no, f"khoảng của `{ma}` phải nằm trong {ts['min']}..{ts['max']} {ts['donVi']} "
                                          "(ngoài khoảng này công thức của mẫu không còn đúng) và nhỏ nhất < lớn nhất")
            if step <= 0 or step > high - low:
                raise ParseError(line_no, f"`buoc` của `{ma}` phải dương và không lớn hơn độ rộng khoảng")
            if not low <= default <= high:
                raise ParseError(line_no, f"`mac-dinh` của `{ma}` phải nằm trong khoảng {low:g}..{high:g}")
            chosen[ma] = {"kieu": "truot", "min": low, "max": high, "buoc": step, "macDinh": default}
    for ma, ts in declared.items():
        chosen.setdefault(ma, {"kieu": "co-dinh", "giaTri": ts["macDinh"]})
    if all(ch["kieu"] == "co-dinh" for ch in chosen.values()):
        raise ParseError(heading_line, "mục `## Tham số` cần ít nhất một tham số thay đổi được (thanh trượt hoặc `chon`)")
    return {ts["ma"]: chosen[ts["ma"]] for ts in khai_bao["thamSo"]}


def _axis(text: str, line_no: int, columns: list[str]) -> Axis:
    for pattern, name in TRANSFORMS:
        match = pattern.match(text.strip())
        if match is not None:
            if match.group(1) not in columns:
                raise ParseError(line_no, f"`do-thi` dùng `{match.group(1)}` nhưng dòng `do:` không có đại lượng này")
            return Axis(match.group(1), name)
    raise ParseError(line_no, f"không hiểu biểu thức `{text.strip()}`; dùng `<mã>`, `<mã>^2`, `1/<mã>`, `ln(<mã>)` "
                              "hoặc `sqrt(<mã>)`")


def parse_experiment(text: str, khai_bao: dict) -> Experiment:
    meta = read_meta(text)
    _, sections, heading_lines = _split(text)
    tham_so = _parameters(sections["## Tham số"], heading_lines["## Tham số"], khai_bao)

    heading = "## Dự đoán"
    pairs = _pairs(sections[heading], ("cau",) + OPTION_KEYS + ("dap-an",), heading)
    du_doan_cau = _required(pairs, "cau", heading_lines[heading], heading)[1]
    lua_chon = [(key, pairs[key][1]) for key in OPTION_KEYS if key in pairs]
    dap_an = None
    if lua_chon:
        if len(lua_chon) < 2 or any(not noi_dung for _, noi_dung in lua_chon):
            raise ParseError(pairs[lua_chon[0][0]][0], "dự đoán trắc nghiệm cần ít nhất hai lựa chọn có nội dung")
        line_no, dap_an = _required(pairs, "dap-an", heading_lines[heading], heading)
        if dap_an not in [key for key, _ in lua_chon]:
            raise ParseError(line_no, "`dap-an` phải là một trong các lựa chọn đã ghi")
    elif "dap-an" in pairs:
        raise ParseError(pairs["dap-an"][0], "`dap-an` chỉ dùng khi dự đoán có các lựa chọn A, B, C, D")

    heading = "## Quan sát"
    pairs = _pairs(sections[heading], ("so-lan-do", "do", "do-thi"), heading)
    line_no, raw = _required(pairs, "so-lan-do", heading_lines[heading], heading)
    if not raw.isdigit() or not MIN_RUNS <= int(raw) <= MAX_RUNS:
        raise ParseError(line_no, f"`so-lan-do` phải là số nguyên từ {MIN_RUNS} đến {MAX_RUNS}")
    so_lan_do = int(raw)
    line_no, raw = _required(pairs, "do", heading_lines[heading], heading)
    cot = [item.strip() for item in raw.split(",") if item.strip()]
    measures = [dl["ma"] for dl in khai_bao["daiLuongDo"]]
    known = list(tham_so) + measures
    unknown = [ma for ma in cot if ma not in known]
    if unknown:
        raise ParseError(line_no, f"mẫu không có `{unknown[0]}`; dùng được: " + ", ".join(known))
    if len(set(cot)) != len(cot) or len(cot) > MAX_COLUMNS:
        raise ParseError(line_no, f"`do` không được lặp và tối đa {MAX_COLUMNS} cột")
    if not any(ma in measures for ma in cot):
        raise ParseError(line_no, "`do` cần ít nhất một đại lượng đo: " + ", ".join(measures))
    do_thi = None
    if "do-thi" in pairs and pairs["do-thi"][1]:
        line_no, raw = pairs["do-thi"]
        parts = raw.split(" theo ")
        if len(parts) != 2:
            raise ParseError(line_no, "`do-thi` cần dạng `<biểu thức trục tung> theo <biểu thức trục hoành>`")
        do_thi = (_axis(parts[0], line_no, cot), _axis(parts[1], line_no, cot))
        for axis in do_thi:
            ts = next((item for item in khai_bao["thamSo"] if item["ma"] == axis.ma), None)
            if ts is not None and ts["kieu"] == "chon":
                raise ParseError(line_no, f"`{axis.ma}` là lựa chọn, không vẽ lên đồ thị được")

    heading = "## Giải thích"
    pairs = _pairs(sections[heading], ("cau", "goi-y-dap-an"), heading)
    giai_thich_cau = _required(pairs, "cau", heading_lines[heading], heading)[1]
    goi_y = _required(pairs, "goi-y-dap-an", heading_lines[heading], heading)[1]

    heading = "## Kết luận"
    if not sections[heading]:
        raise ParseError(heading_lines[heading], "mục `## Kết luận` chưa có nội dung")
    ket_luan = " ".join(line for _, line in sections[heading])

    warnings = []
    slider_columns = [ma for ma in cot if ma in tham_so and tham_so[ma]["kieu"] != "co-dinh"]
    if not slider_columns:
        warnings.append("Bảng số liệu không có tham số nào thay đổi được; học sinh sẽ không thấy quan hệ giữa các đại lượng.")
    return Experiment(meta=meta, tham_so=tham_so, du_doan_cau=du_doan_cau, lua_chon=lua_chon, dap_an=dap_an,
                      so_lan_do=so_lan_do, cot=cot, do_thi=do_thi, giai_thich_cau=giai_thich_cau, goi_y=goi_y,
                      ket_luan=ket_luan, warnings=warnings)
