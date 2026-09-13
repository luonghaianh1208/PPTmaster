#!/usr/bin/env python3
"""Xuất đề kiểm tra KHTN tiếng Anh ra file Word.

Cách dùng:
    python tools/vi/de_thi.py <thư_mục_đề> [--phan tat-ca|de,song-ngu,dap-an] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xuất xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from de_thi_parts import parse  # noqa: E402

SOURCE_NAME = "de.md"
PART_CHOICES = ("de", "song-ngu", "dap-an")
MAX_PATH = 200
FIX_SOURCE = (
    "Viết file de.md trong thư mục đề theo docs/vi/tro-ly/de-khtn-tieng-anh.md rồi chạy lại."
)
FIX_PARSE = "Sửa đúng dòng đó trong de.md theo docs/vi/tro-ly/de-khtn-tieng-anh.md rồi chạy lại."
FIX_DOCX = (
    "Cài thư viện bằng: python -m pip install -r tools/vi/requirements-vi.txt "
    "(hoặc chạy lại CAI-DAT.bat)"
)
FIX_WRITE = "Đóng file Word đang mở rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_INTERNAL = f"Gửi nguyên dòng error.message cho người bảo trì; xem {SOURCE_NAME} có ký tự lạ."


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def result(
    *,
    ready: bool,
    files=(),
    counts: dict | None = None,
    points: float = 0.0,
    warnings=(),
    error: dict | None = None,
) -> dict:
    return {
        "ready": ready,
        "files": [str(path) for path in files],
        "questions": counts or {"part1": 0, "part2": 0, "part3": 0},
        "points": points,
        "warnings": list(warnings),
        "error": error,
    }


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def select_parts(value: str) -> list[str]:
    if value.strip() == "tat-ca":
        return list(PART_CHOICES)
    chosen = [item.strip() for item in value.split(",") if item.strip()]
    unknown = [item for item in chosen if item not in PART_CHOICES]
    if not chosen or unknown:
        raise ValueError(
            "--phan chỉ nhận tat-ca hoặc " + ", ".join(PART_CHOICES)
            + f" (cách nhau bằng dấu phẩy); gặp {value!r}"
        )
    ordered: list[str] = []
    for item in chosen:
        if item not in ordered:
            ordered.append(item)
    return ordered


def load_docx_build():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from de_thi_parts import docx_build

    return docx_build


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = argparse.ArgumentParser(description="Xuất đề kiểm tra KHTN tiếng Anh ra file Word")
    parser.add_argument("folder", type=Path, help="Thư mục đề, chứa file de.md")
    parser.add_argument("--phan", default="tat-ca", help="tat-ca, hoặc de,song-ngu,dap-an")
    parser.add_argument("--plan-only", action="store_true", help="Chỉ kiểm de.md, không ghi file")
    args = parser.parse_args(argv)

    try:
        parts = select_parts(args.phan)
    except ValueError as exc:
        emit(failure("input", str(exc), "Chạy lại với --phan tat-ca"))
        return 1

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục đề: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại de.md bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        exam = parse.parse_exam(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1

    paper_warnings = list(exam.warnings)
    points_warning = exam.points_warning()
    if points_warning:
        paper_warnings.append(points_warning)
    warnings = list(paper_warnings)
    if len(str(folder)) > MAX_PATH:
        warnings.append(
            f"Đường dẫn thư mục đề dài {len(str(folder))} ký tự; Windows có thể không ghi được "
            "file. Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster."
        )

    counts = exam.counts()
    points = exam.total_points()
    if args.plan_only:
        log("Chỉ kiểm de.md, không ghi file.")
        emit(result(ready=True, counts=counts, points=points, warnings=warnings))
        return 0

    try:
        docx_build = load_docx_build()
    except ImportError as exc:
        emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", FIX_DOCX,
                     counts=counts, points=points, warnings=warnings))
        return 1

    for part in parts:
        target = folder / docx_build.FILENAMES[part]
        if target.exists():
            warnings.append(f"Ghi đè file có sẵn: {target.name}")

    try:
        written = docx_build.build(exam, folder, parts, list(paper_warnings))
    except OSError as exc:
        emit(failure("write", f"Không ghi được file Word: {exc}", FIX_WRITE,
                     counts=counts, points=points, warnings=warnings))
        return 1
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến khi dựng file Word: {exc}", FIX_INTERNAL,
                     counts=counts, points=points, warnings=warnings))
        return 1

    log(f"Đã xuất {len(written)} file Word.")
    emit(result(ready=True, files=written, counts=counts, points=points, warnings=warnings))
    return 0


if __name__ == "__main__":
    sys.exit(main())
