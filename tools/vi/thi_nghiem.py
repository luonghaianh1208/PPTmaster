#!/usr/bin/env python3
"""Tạo thí nghiệm ảo (file HTML chạy không cần mạng) và phiếu học tập Word từ thi-nghiem.md.

Cách dùng:
    python tools/vi/thi_nghiem.py <thư_mục_thí_nghiệm> [--phan tat-ca|html|phieu] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import build_html, kiem_so, parse, thu_vien  # noqa: E402

SOURCE_NAME = "thi-nghiem.md"
REVIEW_NAME = "can-soat.md"
PART_CHOICES = ("html", "phieu")
MAX_PATH = 200
GUIDE = "docs/vi/tro-ly/thi-nghiem-ao.md"
MODEL_GUIDE = "docs/vi/tro-ly/mo-hinh-thi-nghiem.md"
FIX_SOURCE = f"Viết file {SOURCE_NAME} trong thư mục thí nghiệm theo {GUIDE} rồi chạy lại."
FIX_PARSE = f"Sửa đúng dòng đó trong {SOURCE_NAME} theo {GUIDE} rồi chạy lại."
FIX_MODEL = f"Sửa mô hình theo khuôn trong {MODEL_GUIDE}; không bỏ công thức, điều kiện áp dụng hay bảng số kiểm."
FIX_CHECK = ("Sửa hàm tinh trong mo-hinh.js cho khớp bảng số kiểm. Chỉ sửa bảng số kiểm khi chính bảng sai, "
             f"và khi đó ghi rõ dòng đã sửa vào mục cần soát để thầy cô kiểm lại ({MODEL_GUIDE}).")
FIX_WRITE = "Đóng file Word hoặc trình duyệt đang mở file cũ rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_INTERNAL = "Gửi nguyên dòng error.message cho người bảo trì."
FIX_ARGS = "Chạy: python tools/vi/thi_nghiem.py <thư_mục_thí_nghiệm> [--phan tat-ca|html|phieu] [--plan-only]"
NEW_MODEL_NOTE = (
    "Mô hình này do AI viết, chưa có người duyệt. Bảng số kiểm cũng do AI tự tính, nên nó chỉ bắt được lỗi lập trình, "
    "không bắt được lỗi hiểu sai kiến thức. Thầy cô kiểm lại công thức dưới đây và tính lại ít nhất hai dòng của bảng "
    "số kiểm bằng máy tính cầm tay trước khi dùng trên lớp."
)


def fix_docx() -> str:
    return (
        f'Cài thư viện bằng: "{sys.executable}" -m pip install -r tools/vi/requirements-vi.txt '
        "(hoặc chạy lại CAI-DAT.bat)"
    )


class ArgumentError(Exception):
    """Tham số dòng lệnh sai; báo bằng JSON thay vì để argparse tự thoát."""


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        raise ArgumentError(message)


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    """In đúng một dòng JSON và không bao giờ ném lỗi."""
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        try:
            sys.stdout.write(text)
        except UnicodeEncodeError:
            buffer = getattr(sys.stdout, "buffer", None)
            if buffer is not None:
                buffer.write(text.encode("utf-8", errors="replace"))
            else:
                sys.stdout.write(json.dumps(payload, ensure_ascii=True) + "\n")
        sys.stdout.flush()
    except OSError:
        pass


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def result(*, ready: bool, files=(), mau: str = "", tham_so=(), so_lan_do: int = 0,
           kiem_so_info: dict | None = None, warnings=(), error: dict | None = None) -> dict:
    return {
        "ready": ready,
        "files": [str(path) for path in files],
        "mau": mau,
        "tham_so": list(tham_so),
        "so_lan_do": so_lan_do,
        "kiem_so": kiem_so_info or {"chay": False, "dat": 0, "tong": 0},
        "warnings": list(warnings),
        "error": error,
    }


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def select_parts(value: str) -> list[str]:
    if value.strip() == "tat-ca":
        return list(PART_CHOICES)
    chosen = [item.strip() for item in value.split(",") if item.strip()]
    if not chosen or any(item not in PART_CHOICES for item in chosen):
        raise ValueError("--phan chỉ nhận tat-ca hoặc " + ", ".join(PART_CHOICES) + f"; gặp {value!r}")
    return [part for part in PART_CHOICES if part in chosen]


def load_phieu():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from thi_nghiem_parts import phieu

    return phieu


def review_text(experiment, model, check: dict | None, warnings: list[str]) -> str:
    formula = model.khai_bao["congThuc"]
    lines = [f"# Cần thầy cô soát — {experiment.meta['tieu-de']}", ""]
    if model.moi:
        lines += [NEW_MODEL_NOTE, ""]
    lines += [
        f"- Mẫu: `{model.ma}` — {model.khai_bao['ten']}",
        f"- Công thức của mô hình: {formula['bieuThuc']}",
        f"- Điều kiện lí tưởng hoá: {formula['dieuKien']}",
    ]
    if formula.get("nguon"):
        lines.append(f"- Nguồn số liệu: {formula['nguon']}")
    if check is None:
        lines.append("- Kiểm số: chưa chạy trên máy này vì không có Node; file HTML tự kiểm mỗi lần mở.")
    else:
        lines.append(f"- Kiểm số: {check['dat']}/{check['tong']} dòng của bảng số kiểm đạt.")
    if model.moi:
        lines += ["", "## Bảng số kiểm do AI viết", ""]
        for index, row in enumerate(model.khai_bao["bangKiem"], 1):
            lines.append(f"{index}. vào {json.dumps(row['vao'], ensure_ascii=False)} → ra "
                         f"{json.dumps(row['ra'], ensure_ascii=False)} (sai số cho phép {row['saiSo']})")
    lines += ["", "## Cảnh báo", ""]
    lines += [f"- {warning}" for warning in warnings] or ["Không có cảnh báo."]
    return "\n".join(lines) + "\n"


def run(args) -> int:
    try:
        parts = select_parts(args.phan)
    except ValueError as exc:
        emit(failure("input", str(exc), "Chạy lại với --phan tat-ca"))
        return 1
    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục thí nghiệm: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}", f"Lưu lại {SOURCE_NAME} bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        meta = parse.read_meta(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1
    try:
        model = thu_vien.load(meta["mau"], folder)
    except thu_vien.ModelError as exc:
        emit(failure("model", str(exc), FIX_MODEL, mau=meta["mau"]))
        return 1
    try:
        experiment = parse.parse_experiment(text, model.khai_bao)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE, mau=model.ma))
        return 1

    summary = {"mau": model.ma, "tham_so": experiment.sliders(), "so_lan_do": experiment.so_lan_do}
    warnings = list(experiment.warnings)
    try:
        check = kiem_so.bang_kiem(model)
    except kiem_so.CheckError as exc:
        emit(failure("check", str(exc), FIX_CHECK, **summary, warnings=warnings))
        return 1
    if check is None:
        kiem_so_info = {"chay": False, "dat": 0, "tong": len(model.khai_bao["bangKiem"])}
        warnings.append("Chưa chạy kiểm số trên máy này vì không có Node; file HTML sẽ tự kiểm khi mở.")
    else:
        kiem_so_info = {"chay": True, "dat": check["dat"], "tong": check["tong"]}
        if check["dat"] != check["tong"]:
            detail = "; ".join(f"dòng {m['dong']} `{m['ma']}`: mong {m['mong']}, được {m['duoc']}" for m in check["truot"])
            emit(failure("check", f"Bảng số kiểm trượt {check['tong'] - check['dat']}/{check['tong']} dòng: {detail}",
                         FIX_CHECK, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
            return 1
    if model.moi:
        warnings.append(f"Mô hình do AI viết, chưa có người duyệt: thầy cô soát công thức và bảng số kiểm trong {REVIEW_NAME}.")
    if len(str(folder)) > MAX_PATH:
        warnings.append(f"Đường dẫn thư mục dài {len(str(folder))} ký tự; Windows có thể không ghi được file. "
                        "Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster.")

    if args.plan_only:
        log(f"Chỉ kiểm {SOURCE_NAME}, không ghi file.")
        emit(result(ready=True, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
        return 0

    phieu = None
    if "phieu" in parts:
        try:
            phieu = load_phieu()
        except ImportError as exc:
            if getattr(exc, "name", None) in ("docx", "lxml"):
                emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", fix_docx(), **summary,
                             kiem_so_info=kiem_so_info, warnings=warnings))
            else:
                emit(failure("internal", f"Lỗi ngoài dự kiến khi nạp bộ dựng Word: {exc}", FIX_INTERNAL, **summary,
                             kiem_so_info=kiem_so_info, warnings=warnings))
            return 1

    files: list[Path] = []
    try:
        if "html" in parts:
            files.append(build_html.write(experiment, model, folder))
        if phieu is not None:
            path, phieu_warnings = phieu.build(experiment, model, folder)
            files.append(path)
            warnings += phieu_warnings
        review = folder / REVIEW_NAME
        review.write_text(review_text(experiment, model, check, warnings), encoding="utf-8", newline="\n")
        files.append(review)
    except build_html.BuildError as exc:
        emit(failure("model", str(exc), FIX_MODEL, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
        return 1
    except OSError as exc:
        emit(failure("write", f"Không ghi được file: {exc}", FIX_WRITE, **summary, kiem_so_info=kiem_so_info,
                     warnings=warnings))
        return 1

    log(f"Đã tạo {len(files)} file.")
    emit(result(ready=True, files=files, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
    return 0


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = JsonArgumentParser(description="Tạo thí nghiệm ảo HTML và phiếu học tập Word")
    parser.add_argument("folder", type=Path, help="Thư mục thí nghiệm, chứa file thi-nghiem.md")
    parser.add_argument("--phan", default="tat-ca", help="tat-ca, hoặc html,phieu")
    parser.add_argument("--plan-only", action="store_true", help="Chỉ kiểm, không ghi file")
    try:
        args = parser.parse_args(argv)
    except ArgumentError as exc:
        emit(failure("input", f"Tham số không hợp lệ: {exc}", FIX_ARGS))
        return 1
    try:
        return run(args)
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến: {exc}", FIX_INTERNAL))
        return 1


if __name__ == "__main__":
    sys.exit(main())
