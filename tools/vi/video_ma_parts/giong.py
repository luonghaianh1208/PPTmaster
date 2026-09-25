"""Giọng đọc từng cảnh: file thầy cô đặt sẵn, giọng máy edge-tts, hoặc bản tạo lần trước còn dùng được."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
from pathlib import Path
from typing import Callable

from video_parts import media

from .lich import GiongInfo

VOICES = {"nu": "vi-VN-HoaiMyNeural", "nam": "vi-VN-NamMinhNeural"}
RATES = {"cham": "-10%", "vua": "+0%", "nhanh": "+15%"}
FIX_GIONG = "Có mạng rồi chạy lại, hoặc đặt sẵn file giọng giong/canh-<số>.mp3 cho từng cảnh."
FIX_EDGE = "Cài edge-tts bằng: python -m pip install -r requirements.txt (ở thư mục gốc repo)."
FIX_FILE = "Xoá hoặc thay file giọng đó rồi chạy lại."


def bam(loi: str, voice: str, rate: str) -> str:
    return hashlib.sha256(f"{voice}|{rate}|{loi}".encode("utf-8")).hexdigest()[:16]


async def _tong_hop(text: str, voice: str, rate: str, out_path: Path) -> list:
    import edge_tts

    moc: list = []
    with open(out_path, "wb") as f:
        communicate = edge_tts.Communicate(text, voice, rate=rate, boundary="SentenceBoundary")
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                f.write(chunk["data"])
            elif chunk["type"] == "SentenceBoundary":
                moc.append(round(chunk["offset"] / 1e7, 3))
    return moc


def tong_hop_edge(text: str, voice: str, rate: str, out_path: Path) -> list:
    try:
        import edge_tts  # noqa: F401
    except ImportError as exc:
        raise media.MediaError("giong", "Chưa cài edge-tts.", FIX_EDGE) from exc
    try:
        return asyncio.run(_tong_hop(text, voice, rate, out_path))
    except Exception as exc:
        raise media.MediaError("giong", f"Không tạo được giọng đọc (thường do mất mạng): {exc}", FIX_GIONG) from exc


def _giay(mp3: Path, do_dai: Callable) -> float:
    if not mp3.is_file() or mp3.stat().st_size == 0:
        raise media.MediaError("giong", f"{mp3.name} rỗng hoặc không có.", FIX_FILE)
    try:
        return float(do_dai(mp3))
    except media.MediaError as exc:
        if str(exc).startswith("ffprobe lỗi"):
            raise media.MediaError("giong", f"{mp3.name} không đọc được (file hỏng?).", FIX_FILE) from exc
        raise


def _so_giong(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def lay_giong(so: int, loi: str, thu_muc: Path, giong: str, toc_do: str,
              tong_hop: Callable = tong_hop_edge, do_dai: Callable = media.probe_duration) -> GiongInfo:
    mp3 = thu_muc / f"canh-{so}.mp3"
    so_giong = thu_muc / f"canh-{so}.json"
    voice, rate = VOICES[giong], RATES[toc_do]
    ma_bam = bam(loi, voice, rate)
    if mp3.is_file():
        ghi = _so_giong(so_giong)
        if not ghi:
            return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=[], uoc_luong=True, nguon="co-san")
        if ghi.get("bam") == ma_bam:
            moc = [float(m) for m in ghi.get("moc", [])]
            return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=moc, uoc_luong=not moc, nguon="may")
    thu_muc.mkdir(parents=True, exist_ok=True)
    tam = mp3.with_name(mp3.name + ".tmp")
    try:
        moc = tong_hop(loi, voice, rate, tam)
        os.replace(tam, mp3)
    except BaseException:
        tam.unlink(missing_ok=True)
        raise
    so_giong.write_text(json.dumps({"bam": ma_bam, "moc": moc}, ensure_ascii=False), encoding="utf-8")
    return GiongInfo(mp3=mp3, giay=_giay(mp3, do_dai), moc_cau=list(moc), uoc_luong=not moc, nguon="may")
