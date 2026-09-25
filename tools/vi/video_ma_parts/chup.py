"""Chụp khung bằng Chromium: mở trang một cảnh, gọi datThoiDiem(t), chụp PNG. `playwright` chỉ import khi dùng."""

from __future__ import annotations

import base64
import contextlib
import os
import sys
from pathlib import Path
from types import SimpleNamespace

from video_parts.media import MediaError

from .phong import CHU_VIET

FIX_CHROMIUM = (
    "Cài Chromium bằng: powershell -NoProfile -ExecutionPolicy Bypass -File tools\\vi\\pptmaster.ps1 "
    "-Action tool -Name chromium"
)
FIX_DUNG = "Chạy lại một lần; vẫn lỗi thì dán nguyên thông báo này cho người bảo trì."
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
    # Nền cảnh trước và ảnh phải giải mã xong trước khung đầu, nếu không khung t = 0 có thể trống.
    page.evaluate("() => Promise.all(Array.from(document.images).map((i) => i.decode().catch(() => null)))")


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


def so_tien_trinh() -> int:
    return max(1, min(4, (os.cpu_count() or 2) // 2))


def chia_dai(so_khung_moi_canh: list, so_tien_trinh: int) -> list:
    """Chia các cảnh thành dải liên tiếp [dau, cuoi) có tổng khung gần đều (tham lam theo tổng tích luỹ)."""
    n = len(so_khung_moi_canh)
    so_dai = max(1, min(so_tien_trinh, n))
    tong = sum(so_khung_moi_canh)
    dai, dau, tich_luy = [], 0, 0
    for k in range(n - 1):
        if len(dai) == so_dai - 1:
            break
        tich_luy += so_khung_moi_canh[k]
        dich = tong * (len(dai) + 1) / so_dai
        vuot_neu_them = tich_luy + so_khung_moi_canh[k + 1] - dich
        if tich_luy >= dich or vuot_neu_them > dich - tich_luy:
            dai.append((dau, k + 1))
            dau = k + 1
    dai.append((dau, n))
    return dai


def _data_url(png: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii")


def chup_dai(cong_viec: dict) -> int:
    """Chụp một dải cảnh trong một Chromium riêng. Chạy được trong tiến trình con Windows `spawn`."""
    tools_vi = str(Path(__file__).resolve().parents[1])
    if tools_vi not in sys.path:
        sys.path.insert(0, tools_vi)
    from . import trang

    cac_du = [dict(du) for du in cong_viec["cac_du"]]
    models = {int(so): SimpleNamespace(js=js) for so, js in cong_viec["models_js"].items()}
    dau, cuoi, fps = cong_viec["dau"], cong_viec["cuoi"], cong_viec["fps"]
    khung_dau, so_khung = cong_viec["khung_dau"], cong_viec["so_khung"]
    thu_muc = Path(cong_viec["thu_muc_anh"])

    def html(k: int) -> str:
        return trang.dung_trang(cac_du[k], models.get(cac_du[k]["so"]))

    def can_nen(k: int) -> bool:
        return k < len(cac_du) and bool(cac_du[k]["co"].get("lauBang"))

    da_ghi = 0
    with trinh_duyet() as browser:
        page = trang_moi(browser)
        if dau > 0 and can_nen(dau):
            # Khung cuối cảnh trước (t = (số khung − 1)/fps) dựng lại tại chỗ, không chờ tiến trình khác.
            mo_trang(page, html(dau - 1))
            page.evaluate("(t) => window.datThoiDiem(t)", (so_khung[dau - 1] - 1) / fps)
            cac_du[dau]["nenTruoc"] = _data_url(page.screenshot(type="png"))
        for k in range(dau, cuoi):
            print(f"Chụp cảnh {cac_du[k]['so']} ({so_khung[k]} khung)...", file=sys.stderr, flush=True)
            chup_canh(page, html(k), so_khung[k], fps, thu_muc, khung_dau[k])
            da_ghi += so_khung[k]
            if k + 1 < cuoi and can_nen(k + 1):
                cuoi_k = thu_muc / f"f{khung_dau[k] + so_khung[k] - 1:06d}.png"
                cac_du[k + 1]["nenTruoc"] = _data_url(cuoi_k.read_bytes())
    return da_ghi


def _chup_dai_con(cong_viec: dict) -> int:
    """Vỏ cho tiến trình con: MediaError không gửi ngược qua pickle được nên đổi sang RuntimeError."""
    try:
        return chup_dai(cong_viec)
    except MediaError as exc:
        raise RuntimeError(f"{exc.step}: {exc.message}") from None


def _loi_dai(cac_du: list, dau: int, cuoi: int, exc: BaseException) -> MediaError:
    a, b = cac_du[dau]["so"], cac_du[cuoi - 1]["so"]
    canh = f"cảnh {a}" if a == b else f"cảnh {a}–{b}"
    return MediaError("dung", f"Chụp khung lỗi ở {canh}: {type(exc).__name__}: {exc}", FIX_DUNG)


def chup_song_song(cac_du: list, models_js: dict, so_khung: list, fps: int, thu_muc_anh, so_tt: int) -> None:
    khung_dau = [0]
    for n in so_khung[:-1]:
        khung_dau.append(khung_dau[-1] + n)
    Path(thu_muc_anh).mkdir(parents=True, exist_ok=True)
    viec = [{"cac_du": cac_du, "models_js": models_js, "dau": a, "cuoi": b, "khung_dau": khung_dau,
             "so_khung": list(so_khung), "fps": fps, "thu_muc_anh": str(thu_muc_anh)}
            for a, b in chia_dai(so_khung, so_tt)]
    if len(viec) == 1:
        try:
            chup_dai(viec[0])
        except MediaError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _loi_dai(cac_du, viec[0]["dau"], viec[0]["cuoi"], exc) from exc
        return
    from concurrent.futures import ProcessPoolExecutor

    with ProcessPoolExecutor(max_workers=len(viec)) as pool:
        cac_tuong_lai = [(v, pool.submit(_chup_dai_con, v)) for v in viec]
        loi = None
        for v, tl in cac_tuong_lai:
            try:
                tl.result()
            except Exception as exc:  # noqa: BLE001
                loi = loi or _loi_dai(cac_du, v["dau"], v["cuoi"], exc)
    if loi is not None:
        raise loi
