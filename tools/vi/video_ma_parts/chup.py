"""Chụp khung bằng Chromium: mở trang một cảnh, gọi datThoiDiem(t), chụp PNG. `playwright` chỉ import khi dùng."""

from __future__ import annotations

import contextlib
import os
from pathlib import Path

from video_parts.media import MediaError

from .phong import CHU_VIET

FIX_CHROMIUM = (
    "Cài Chromium bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 "
    "-Action tool -Name chromium"
)
VIEWPORT = {"width": 1280, "height": 720}


def _mo(p):
    try:
        return p.chromium.launch(args=["--no-sandbox"])
    except Exception as first:
        base = Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright"
        for pattern in ("chromium_headless_shell-*/*/chrome-headless-shell.exe", "chromium-*/*/chrome.exe"):
            for exe in sorted(base.glob(pattern), reverse=True):
                try:
                    return p.chromium.launch(executable_path=str(exe), args=["--no-sandbox"])
                except Exception:
                    continue
        raise MediaError("chromium", f"Không mở được Chromium: {first}", FIX_CHROMIUM) from first


@contextlib.contextmanager
def trinh_duyet():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise MediaError("chromium", "Chưa cài playwright (Chromium).", FIX_CHROMIUM) from exc
    with sync_playwright() as p:
        browser = _mo(p)
        try:
            yield browser
        finally:
            browser.close()


def trang_moi(browser):
    return browser.new_page(viewport=VIEWPORT, device_scale_factor=1)


def mo_trang(page, html: str) -> None:
    page.set_content(html)
    page.wait_for_function("window.THI_VIDEO && window.THI_VIDEO.san === true", timeout=30000)
    page.evaluate(
        "(chuViet) => document.fonts.load(\"40px 'Itim'\", document.body.innerText + chuViet)", CHU_VIET,
    )
    page.evaluate("() => document.fonts.ready.then(() => true)")


def kiem_tran(page, html: str) -> list:
    mo_trang(page, html)
    return list(page.evaluate("() => window.THI_VIDEO.kiemTran()"))


def chup_canh(page, html: str, so_khung: int, fps: int, thu_muc: Path, so_dau: int, ghi_log=None) -> int:
    mo_trang(page, html)
    thu_muc.mkdir(parents=True, exist_ok=True)
    for i in range(so_khung):
        page.evaluate("(t) => window.datThoiDiem(t)", i / fps)
        page.screenshot(path=str(thu_muc / f"f{so_dau + i:06d}.png"), type="png")
        if ghi_log is not None and (i + 1) % 60 == 0:
            ghi_log(f"  đã chụp {i + 1}/{so_khung} khung của cảnh này")
    return so_dau + so_khung


def chup_cuoi(page, html: str, duong_dan: Path) -> None:
    mo_trang(page, html)
    page.evaluate("() => window.datThoiDiem(window.THI_VIDEO.thoiDiemCuoi())")
    duong_dan.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(duong_dan), type="png")
