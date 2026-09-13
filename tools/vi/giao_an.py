#!/usr/bin/env python3
"""Soạn Kế hoạch bài dạy (giáo án) tích hợp năng lực số và năng lực AI.

Cách dùng:
    python tools/vi/giao_an.py xuat <thư_mục_giáo_án> [--plan-only]
    python tools/vi/giao_an.py trich-sgk <file_sgk.md> --bai "<tên bài>" [--ra <file.md>]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from giao_an_parts import frameworks, parse, sgk  # noqa: E402

SOURCE_NAME = "giao-an.md"
SLICE_NAME = "sgk-trich.md"
MAX_PATH = 200
FIX_SOURCE = (
    "Viết file giao-an.md trong thư mục giáo án theo docs/vi/tro-ly/giao-an.md rồi chạy lại."
)
FIX_PARSE = "Sửa đúng dòng đó trong giao-an.md theo docs/vi/tro-ly/giao-an.md rồi chạy lại."


def fix_docx() -> str:
    """Lệnh cài dùng đúng Python đang chạy công cụ (có thể là Python của venv)."""
    return (
        f'Cài thư viện bằng: "{sys.executable}" -m pip install -r tools/vi/requirements-vi.txt '
        "(hoặc chạy lại CAI-DAT.bat)"
    )


FIX_WRITE = "Đóng file Word đang mở rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_WRITE_SLICE = "Kiểm tra thư mục của --ra có tồn tại và ổ đĩa còn trống, rồi chạy lại."
FIX_INTERNAL = "Gửi nguyên dòng error.message cho người bảo trì."
FIX_ARGS = (
    "Chạy: python tools/vi/giao_an.py xuat <thư_mục_giáo_án> [--plan-only], hoặc "
    "python tools/vi/giao_an.py trich-sgk <file_sgk.md> --bai \"<tên bài>\" [--ra <file.md>]"
)


class ArgumentError(Exception):
    """Tham số dòng lệnh sai; báo bằng JSON thay vì để argparse tự thoát."""


class JsonArgumentParser(argparse.ArgumentParser):
    """Lệnh con tạo bằng add_subparsers kế thừa lớp này, nên cũng báo lỗi bằng ngoại lệ."""

    def error(self, message):
        raise ArgumentError(message)


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    """In đúng một dòng JSON và không bao giờ ném lỗi, để main không phải trả lời lần thứ hai."""
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
        # stdout đã đóng hoặc hỏng (ví dụ bên đọc thoát sớm): không còn cách nào trả lời thêm.
        pass


def result(*, ready: bool, files=(), warnings=(), error: dict | None = None, **extra) -> dict:
    payload = {
        "ready": ready,
        "files": [str(path) for path in files],
        "warnings": list(warnings),
        "error": error,
    }
    payload.update(extra)
    return payload


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def load_docx_build():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from giao_an_parts import docx_build

    return docx_build


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _lesson_extra(lesson) -> dict:
    return {
        "activities": len(lesson.activities),
        "periods": lesson.periods,
        "minutes": lesson.minutes,
        "nls_codes": lesson.nls_codes,
        "ai_codes": lesson.ai_codes,
    }


def command_export(args) -> int:
    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục giáo án: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại giao-an.md bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        lesson = parse.parse_lesson(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1

    try:
        frameworks.validate(lesson, frameworks.load_file())
    except frameworks.FrameworkError as exc:
        emit(failure("framework", exc.message, exc.fix, **_lesson_extra(lesson)))
        return 1

    extra = _lesson_extra(lesson)
    paper_warnings = list(lesson.warnings)
    warnings = list(paper_warnings)
    if len(str(folder)) > MAX_PATH:
        warnings.append(
            f"Đường dẫn thư mục giáo án dài {len(str(folder))} ký tự; Windows có thể không ghi "
            "được file. Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster."
        )

    if args.plan_only:
        log("Chỉ kiểm giao-an.md, không ghi file.")
        emit(result(ready=True, warnings=warnings, **extra))
        return 0

    try:
        docx_build = load_docx_build()
    except ImportError as exc:
        if getattr(exc, "name", None) not in ("docx", "lxml"):
            emit(failure("internal", f"Lỗi ngoài dự kiến khi nạp bộ dựng Word: {exc}", FIX_INTERNAL,
                         warnings=warnings, **extra))
            return 1
        emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", fix_docx(),
                     warnings=warnings, **extra))
        return 1

    for name in (docx_build.FILENAME, docx_build.REVIEW_FILENAME):
        if (folder / name).exists():
            warnings.append(f"Ghi đè file có sẵn: {name}")

    try:
        plan_path = docx_build.build(lesson, folder)
        review_path = docx_build.write_review(lesson, folder, paper_warnings)
    except OSError as exc:
        emit(failure("write", f"Không ghi được file: {exc}", FIX_WRITE,
                     warnings=warnings, **extra))
        return 1
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến khi dựng file Word: {exc}", FIX_INTERNAL,
                     warnings=warnings, **extra))
        return 1

    log("Đã xuất giáo án và file ghi chú.")
    emit(result(ready=True, files=[plan_path, review_path], warnings=warnings, **extra))
    return 0


def command_cut(args) -> int:
    source = args.sgk.expanduser().resolve()
    if not source.is_file():
        emit(failure("input", f"Không có file SGK: {source}",
                     "Chuyển SGK PDF sang Markdown một lần bằng "
                     "skills/ppt-master/scripts/source_to_md.py rồi chạy lại."))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}",
                     "Lưu lại file SGK bằng bảng mã UTF-8 rồi chạy lại"))
        return 1
    try:
        heading, body, warnings = sgk.extract(text, args.bai)
    except sgk.SgkError as exc:
        emit(failure("parse", exc.message, exc.fix))
        return 1
    target = (args.ra or source.parent / SLICE_NAME).expanduser().resolve()
    if target == source:
        emit(failure("input", f"--ra trùng với file SGK nguồn: {target}",
                     "Chọn một đường dẫn --ra khác, để không ghi đè file SGK đã chuyển đổi."))
        return 1
    try:
        target.write_text(body, encoding="utf-8")
    except OSError as exc:
        emit(failure("write", f"Không ghi được {target}: {exc}", FIX_WRITE_SLICE))
        return 1
    size_kb = round(len(body.encode("utf-8")) / 1024, 1)
    log(f"Đã cắt phần {heading!r} ra {target}.")
    emit(result(ready=True, files=[target], warnings=warnings, heading=heading, size_kb=size_kb))
    return 0


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = JsonArgumentParser(description="Soạn giáo án tích hợp năng lực số và năng lực AI")
    commands = parser.add_subparsers(dest="command", required=True)

    export = commands.add_parser("xuat", help="Xuất giáo án từ giao-an.md")
    export.add_argument("folder", type=Path, help="Thư mục giáo án, chứa file giao-an.md")
    export.add_argument("--plan-only", action="store_true", help="Chỉ kiểm, không ghi file")
    export.set_defaults(handler=command_export)

    cut = commands.add_parser("trich-sgk", help="Cắt phần một bài ra khỏi SGK đã chuyển Markdown")
    cut.add_argument("sgk", type=Path, help="File SGK dạng Markdown")
    cut.add_argument("--bai", required=True, help="Tên hoặc số bài, ví dụ \"Bài 5\"")
    cut.add_argument("--ra", type=Path, default=None, help="Đường dẫn file kết quả")
    cut.set_defaults(handler=command_cut)

    try:
        args = parser.parse_args(argv)
    except ArgumentError as exc:
        emit(failure("input", f"Tham số không hợp lệ: {exc}", FIX_ARGS))
        return 1
    try:
        return args.handler(args)
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến: {exc}", FIX_INTERNAL))
        return 1


if __name__ == "__main__":
    sys.exit(main())
