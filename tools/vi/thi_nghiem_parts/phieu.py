"""Dựng phiếu học tập Word cho một thí nghiệm ảo: trang học sinh và trang giáo viên."""

from __future__ import annotations

import math
from pathlib import Path

from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm, Pt

from word_parts import base

from . import kiem_so, tham_chieu

FILENAME = "phieu-hoc-tap.docx"
DOTS = "." * 118
GRID_ROWS, GRID_COLUMNS = 14, 20
GRID_CELL_CM = 0.75
TRANSFORM_LABELS = {"binh-phuong": "({})^2^", "nghich-dao": "1/({})", "ln": "ln({})", "can": "√({})", "khong": "{}"}


def label(model, ma: str) -> str:
    item = model.tham_so(ma) or model.dai_luong(ma)
    return item["ten"] + (f" ({item['donVi']})" if item.get("donVi") else "")


def axis_label(model, axis) -> str:
    return TRANSFORM_LABELS[axis.phep].format(label(model, axis.ma))


def _format(value, digits: int) -> str:
    if value is None:
        return "—"
    if isinstance(value, str):
        return value
    return f"{value:.{digits}f}".replace(".", ",")


def _digits(model, ma: str) -> int:
    measure = model.dai_luong(ma)
    if measure is not None:
        return measure["chuSo"]
    step = str(model.tham_so(ma).get("buoc", 1))
    return len(step.split(".")[1]) if "." in step else 0


def _cell(model, ma: str, point: dict, result: dict) -> str:
    parameter = model.tham_so(ma)
    if parameter is None:
        return _format(result.get(ma), _digits(model, ma))
    if parameter["kieu"] == "chon":
        return next(option["ten"] for option in parameter["luaChon"] if option["ma"] == point[ma])
    return _format(point[ma], _digits(model, ma))


def ideal_table(experiment, model) -> tuple[list[list[str]] | None, str]:
    """Bảng số liệu lí tưởng cho giáo viên. Trả về (bảng hoặc None, cảnh báo hoặc chuỗi rỗng)."""
    varied = next((ma for ma in experiment.cot
                   if ma in experiment.tham_so and experiment.tham_so[ma]["kieu"] == "truot"), None)
    if varied is None:
        return None, "Phiếu không có bảng số liệu lí tưởng vì bảng đo không có tham số dạng thanh trượt."
    chosen = experiment.tham_so[varied]
    count = experiment.so_lan_do
    steps = round((chosen["max"] - chosen["min"]) / chosen["buoc"])
    # floor(x + 0.5) thay cho round(): round() của Python làm tròn về số chẵn nên các mốc bị lệch về một phía.
    values = sorted({round(chosen["min"] + math.floor(index * steps / (count - 1) + 0.5) * chosen["buoc"], 10)
                     for index in range(count)})
    points = []
    for value in values:
        point = {ma: (ch["giaTri"] if ch["kieu"] == "co-dinh" else ch["macDinh"]) for ma, ch in experiment.tham_so.items()}
        point[varied] = value
        points.append(point)
    if model.ma in tham_chieu.THAM_CHIEU:
        results = [tham_chieu.THAM_CHIEU[model.ma](point) for point in points]
    else:
        try:
            results = kiem_so.luoi(model, points)
        except kiem_so.CheckError as exc:
            return None, f"Phiếu không có bảng số liệu lí tưởng: {exc}"
        if results is None:
            return None, "Phiếu không có bảng số liệu lí tưởng vì máy này không có Node để chạy mô hình mới."
    return [[_cell(model, ma, point, result) for ma in experiment.cot] for point, result in zip(points, results)], ""


def _heading(document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(8)
    base.write(paragraph, text, bold=True)


def _line(document, text: str, **kwargs) -> None:
    base.write(document.add_paragraph(), text, **kwargs)


def _dots(document, count: int) -> None:
    for _ in range(count):
        document.add_paragraph(DOTS)


def _data_table(document, headers: list[str], rows: list[list[str]]) -> None:
    table = document.add_table(rows=1 + len(rows), cols=len(headers))
    base.grid_borders(table)
    for index, header in enumerate(headers):
        base.fill_cell(table.rows[0].cells[index], [header])
    for row_index, row in enumerate(rows, 1):
        for column, value in enumerate(row):
            base.fill_cell(table.rows[row_index].cells[column], [value], bold_first=False)
        table.rows[row_index].height = Cm(0.9)


def _grid(document) -> None:
    table = document.add_table(rows=GRID_ROWS, cols=GRID_COLUMNS)
    base.grid_borders(table)
    base.set_widths(table, [Cm(GRID_CELL_CM)] * GRID_COLUMNS)
    for row in table.rows:
        row.height = Cm(GRID_CELL_CM)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY


def build(experiment, model, folder: Path) -> tuple[Path, list[str]]:
    warnings: list[str] = []
    document = base.new_document(margins_cm=(1.8, 1.8, 2.0, 1.5), size_pt=13, line_spacing=1.2)
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(title, "PHIẾU HỌC TẬP", bold=True)
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(subtitle, f"{experiment.meta['tieu-de']} — {experiment.meta['mon']} {experiment.meta['lop']}", bold=True)
    _line(document, "Họ và tên / Nhóm: ................................................................  Lớp: ..................")

    _heading(document, "1. Dự đoán (làm trước khi chạy thí nghiệm)")
    _line(document, experiment.du_doan_cau)
    for ma, noi_dung in experiment.lua_chon:
        _line(document, f"**{ma}.** {noi_dung}")
    _line(document, "Dự đoán của em: " + "." * 90)

    _heading(document, "2. Quan sát")
    _line(document, f"Thay đổi tham số, làm thí nghiệm và ghi {experiment.so_lan_do} lần đo vào bảng.")
    headers = ["Lần"] + [label(model, ma) for ma in experiment.cot]
    _data_table(document, headers, [[str(index)] + [""] * len(experiment.cot) for index in range(1, experiment.so_lan_do + 1)])
    if experiment.do_thi is not None:
        tung, hoanh = experiment.do_thi
        _line(document, f"Vẽ đồ thị. Trục tung: {axis_label(model, tung)}. Trục hoành: {axis_label(model, hoanh)}.")
        _grid(document)

    _heading(document, "3. Giải thích")
    _line(document, experiment.giai_thich_cau)
    _dots(document, 4)
    _heading(document, "Kết luận")
    _dots(document, 3)

    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    teacher = document.add_paragraph()
    teacher.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(teacher, "DÀNH CHO GIÁO VIÊN — KHÔNG PHÁT CHO HỌC SINH", bold=True)
    if experiment.dap_an:
        _line(document, f"**Đáp án phần dự đoán:** {experiment.dap_an}")
    _line(document, f"**Gợi ý phần giải thích:** {experiment.goi_y}")
    _line(document, f"**Kết luận:** {experiment.ket_luan}")
    formula = model.khai_bao["congThuc"]
    _line(document, f"**Mô hình của thí nghiệm ảo:** {formula['bieuThuc']}")
    _line(document, f"**Điều kiện lí tưởng hoá:** {formula['dieuKien']}")
    if formula.get("nguon"):
        _line(document, f"**Nguồn số liệu:** {formula['nguon']}")
    rows, warning = ideal_table(experiment, model)
    if rows is None:
        warnings.append(warning)
    else:
        _heading(document, "Số liệu lí tưởng (không có sai số đo)")
        _data_table(document, [label(model, ma) for ma in experiment.cot], rows)

    path = folder / FILENAME
    document.save(str(path))
    return path, warnings
