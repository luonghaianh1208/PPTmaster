#!/usr/bin/env python3
"""Dựng video giải thích kiểu viết tay từ video.md (giọng đọc, cảnh vẽ bằng mã, phụ đề).

  python tools/vi/video_ma.py <thư_mục> [--plan-only] [--xem-truoc]

stdout đúng một dòng JSON. Hướng dẫn: docs/vi/tro-ly/video-giai-thich.md
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import thu_vien  # noqa: E402
from video_ma_parts import chup, ghep, giong, kiem, lich, parse, trang  # noqa: E402
from video_parts import media  # noqa: E402

FIX_INPUT = "Viết video.md trong thư mục dự án (xem docs/vi/tro-ly/video-giai-thich.md) rồi chạy lại."
FIX_INTERNAL = "Lỗi ngoài dự kiến; dán nguyên thông báo này cho người bảo trì."


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))
    sys.stdout.flush()


def co_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None


def co_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return False
    base = Path(root) / "ms-playwright"
    if not (base.is_dir() and (any(base.glob("chromium-*")) or any(base.glob("chromium_headless_shell-*")))):
        return False
    return importlib.util.find_spec("playwright") is not None


def _mo_hinh(video: parse.Video, thu_muc: Path) -> dict:
    return {c.so: thu_vien.load(c.truong["mau"][0], thu_muc) for c in video.canh if c.loai == "thi-nghiem"}


def _trang(video, cac_lich, models) -> list:
    return [trang.dung_trang(lich.du_lieu_canh(c, cl, models.get(c.so)), models.get(c.so)) for c, cl in zip(video.canh, cac_lich)]


def _kiem_tran_tat_ca(page, video, trang_html) -> None:
    for canh, html in zip(video.canh, trang_html):
        tran = chup.kiem_tran(page, html)
        if tran:
            raise kiem.CanhError(canh.so, f"chữ ở mục `{', '.join(tran)}` tràn khung. Rút ngắn nội dung hoặc chia thành hai cảnh.")


def _xem_truoc(video: parse.Video, thu_muc: Path, warnings: list) -> dict:
    gia = [lich.GiongInfo(mp3=None, giay=8.0, moc_cau=[], uoc_luong=True, nguon="may") for _ in video.canh]
    cac_lich, _ = lich.dung_lich(video.canh, gia, kiem_moc=False)
    trang_html = _trang(video, cac_lich, _mo_hinh(video, thu_muc))
    ra = thu_muc / "xem-truoc"
    files = []
    with chup.trinh_duyet() as browser:
        page = chup.trang_moi(browser)
        _kiem_tran_tat_ca(page, video, trang_html)
        for canh, html in zip(video.canh, trang_html):
            chup.chup_cuoi(page, html, ra / f"canh-{canh.so}.png")
            files.append(f"xem-truoc/canh-{canh.so}.png")
    return {"files": files, "so_canh": len(video.canh), "thoi_luong_giay": None,
            "phong_cach": video.meta["phong-cach"], "giong": None}


def _dung(video: parse.Video, thu_muc: Path, warnings: list) -> dict:
    if not co_ffmpeg():
        raise media.MediaError("ffmpeg", "Chưa có FFmpeg.", media.FIX_FFMPEG)
    if not co_chromium():
        raise media.MediaError("chromium", "Chưa cài Chromium hoặc playwright.", chup.FIX_CHROMIUM)
    cac_giong = [giong.lay_giong(c.so, c.loi, thu_muc / "giong", video.meta["giong"], video.meta["toc-do"]) for c in video.canh]
    cac_lich, canh_bao = lich.dung_lich(video.canh, cac_giong)
    warnings.extend(canh_bao)
    trang_html = _trang(video, cac_lich, _mo_hinh(video, thu_muc))
    lam = thu_muc / ".khung"
    shutil.rmtree(lam, ignore_errors=True)
    anh = lam / "anh"
    anh.mkdir(parents=True)
    try:
        with chup.trinh_duyet() as browser:
            page = chup.trang_moi(browser)
            _kiem_tran_tat_ca(page, video, trang_html)
            so = 0
            for canh, cl, html in zip(video.canh, cac_lich, trang_html):
                log(f"Chụp cảnh {canh.so}/{len(video.canh)} ({cl.so_khung} khung)...")
                so = chup.chup_canh(page, html, cl.so_khung, lich.FPS, anh, so, log)
        log("Ghép video bằng FFmpeg...")
        files = ghep.ghep_video(thu_muc, cac_lich, cac_giong, video.meta["phu-de"])
    finally:
        shutil.rmtree(lam, ignore_errors=True)
    nguon = {g.nguon for g in cac_giong}
    return {"files": files, "so_canh": len(video.canh), "thoi_luong_giay": round(sum(cl.thoi_luong for cl in cac_lich), 2),
            "phong_cach": video.meta["phong-cach"], "giong": nguon.pop() if len(nguon) == 1 else "hon-hop"}


def chay(thu_muc: Path, plan_only: bool, xem_truoc: bool, warnings: list) -> dict:
    md = thu_muc / "video.md"
    if not thu_muc.is_dir() or not md.is_file():
        raise media.MediaError("input", f"Không thấy {md}.", FIX_INPUT)
    video = parse.parse(md.read_text(encoding="utf-8-sig"))
    warnings.extend(kiem.kiem(video, thu_muc))
    if plan_only:
        return {"files": [], "so_canh": len(video.canh), "thoi_luong_giay": None,
                "phong_cach": video.meta["phong-cach"], "giong": None}
    if xem_truoc:
        if not co_chromium():
            raise media.MediaError("chromium", "Chưa cài Chromium hoặc playwright.", chup.FIX_CHROMIUM)
        return _xem_truoc(video, thu_muc, warnings)
    return _dung(video, thu_muc, warnings)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Dựng video giải thích kiểu viết tay từ video.md", add_help=False)
    ap.add_argument("thu_muc")
    ap.add_argument("--plan-only", action="store_true")
    ap.add_argument("--xem-truoc", action="store_true")
    try:
        args = ap.parse_args(argv)
    except SystemExit:
        emit({"ready": False, "files": [], "so_canh": 0, "thoi_luong_giay": None, "phong_cach": None, "giong": None, "warnings": [],
              "error": {"step": "input", "message": "Sai tham số dòng lệnh.", "fix": "Dùng: python tools/vi/video_ma.py <thư_mục> [--plan-only] [--xem-truoc]"}})
        return 1
    warnings: list = []
    base = {"ready": False, "files": [], "so_canh": 0, "thoi_luong_giay": None, "phong_cach": None, "giong": None}
    try:
        kq = chay(Path(args.thu_muc).resolve(), args.plan_only, args.xem_truoc, warnings)
        emit({**base, **kq, "ready": True, "warnings": warnings, "error": None})
        return 0
    except parse.ParseError as exc:
        error = {"step": "parse", "message": str(exc), "fix": "Sửa đúng dòng đó trong video.md rồi chạy lại."}
    except kiem.CanhError as exc:
        error = {"step": "canh", "message": str(exc), "fix": "Rút gọn hoặc sửa nội dung cảnh đó theo thông báo."}
    except media.MediaError as exc:
        error = {"step": exc.step, "message": str(exc), "fix": exc.fix}
    except OSError as exc:
        error = {"step": "write", "message": f"Không ghi được file: {exc}", "fix": "Đóng file đang mở và kiểm tra ổ đĩa rồi chạy lại."}
    except Exception as exc:  # noqa: BLE001
        error = {"step": "internal", "message": f"{type(exc).__name__}: {exc}", "fix": FIX_INTERNAL}
    emit({**base, "warnings": warnings, "error": error})
    return 1


if __name__ == "__main__":
    sys.exit(main())
