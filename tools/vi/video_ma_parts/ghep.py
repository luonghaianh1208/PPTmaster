"""Ghép khung hình + tiếng + phụ đề thành video.mp4 bằng FFmpeg. Thư mục dự án chỉ được truyền dưới dạng cwd và đường dẫn tuyệt đối."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path

from video_parts import media, srt

from . import karaoke
from .lich import DAN_DAU, FPS
from .phong import FONT as ITIM_FONT, TEN as ITIM_TEN

_MARKUP_RE = re.compile(r"\*\*|~|\^")
BIEN_DO_NHIEU = 0.002
HAT_NHIEU = 1234
STYLE = f"FontName={ITIM_TEN},FontSize=16,Outline=1.5,Shadow=0,Spacing=0.5,MarginV=22"
FONTS_REL = ".khung/fonts"


def cues_phu_de(cac_lich: list) -> list:
    cues = []
    so = 0
    for cl in cac_lich:
        het = cl.bat_dau + DAN_DAU + cl.giay_giong
        for k, cau in enumerate(cl.cau):
            start = cl.bat_dau + cl.moc_cau[k]
            end = cl.bat_dau + cl.moc_cau[k + 1] if k + 1 < len(cl.cau) else het
            so += 1
            cues.append(srt.Cue(index=so, start=start, end=max(end, start + 0.1), text=_MARKUP_RE.sub("", cau)))
    return cues


def lenh_am_canh(mp3: Path, wav: Path, thoi_luong: float) -> list:
    # Lớp nhiễu hồng rất nhỏ: loa Bluetooth/HDMI tự tắt khi gặp im lặng tuyệt đối và nuốt âm đầu câu sau.
    nhieu = f"anoisesrc=d={thoi_luong:.3f}:c=pink:r=44100:a={BIEN_DO_NHIEU}:seed={HAT_NHIEU}"
    loc = (
        f"[0:a]aresample=44100,aformat=channel_layouts=mono,adelay={int(round(DAN_DAU * 1000))}:all=1,"
        f"apad=whole_dur={thoi_luong:.3f}[g];[1:a]aformat=channel_layouts=mono[n];"
        "[g][n]amix=inputs=2:duration=first:normalize=0[a]"
    )
    return [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(mp3),
        "-f", "lavfi", "-i", nhieu, "-filter_complex", loc, "-map", "[a]",
        "-t", f"{thoi_luong:.3f}", "-ar", "44100", "-ac", "1", "-c:a", "pcm_s16le", str(wav),
    ]


def lenh_video(danh_sach_am: Path, out_mp4: Path, fps: int, phu_de_tuong_doi) -> list:
    cmd = [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-framerate", str(fps), "-i", ".khung/anh/f%06d.png",
        "-f", "concat", "-safe", "0", "-i", str(danh_sach_am),
    ]
    if phu_de_tuong_doi:
        if str(phu_de_tuong_doi).endswith(".ass"):
            cmd += ["-vf", f"subtitles={phu_de_tuong_doi}:fontsdir={FONTS_REL}"]
        else:
            cmd += ["-vf", f"subtitles={phu_de_tuong_doi}:fontsdir={FONTS_REL}:force_style='{STYLE}'"]
    cmd += ["-r", str(fps), "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart", str(out_mp4)]
    return cmd


def _chay(cmd: list, run, cwd: Path) -> None:
    try:
        proc = run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=7200, cwd=cwd)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise media.MediaError("ffmpeg", f"Không chạy được FFmpeg: {exc}", media.FIX_FFMPEG) from exc
    if proc.returncode != 0:
        loi = (proc.stderr or "").strip()[-400:]
        raise media.MediaError("dung", f"FFmpeg lỗi (mã {proc.returncode}): {loi}", "Báo nội dung lỗi cho người bảo trì.")


def ghep_video(thu_muc: Path, cac_lich: list, cac_giong: list, phu_de: str, fps: int = FPS, run=subprocess.run) -> list:
    lam = thu_muc / ".khung"
    wavs = []
    for cl, giong in zip(cac_lich, cac_giong):
        wav = lam / f"am-{cl.so}.wav"
        _chay(lenh_am_canh(giong.mp3, wav, cl.thoi_luong), run, thu_muc)
        wavs.append(wav)
    danh_sach = lam / "am.txt"
    danh_sach.write_text(media.build_audio_concat_text(wavs), encoding="utf-8")
    cac_cue = cues_phu_de(cac_lich)
    cues = srt.render_srt(cac_cue)
    files = ["video.mp4"]
    burn = None
    if phu_de == "hinh":
        thoat = [srt.Cue(index=c.index, start=c.start, end=c.end, text=c.text.replace("{", "\\{").replace("}", "\\}"))
                 for c in cac_cue]
        (lam / "phu-de.srt").write_text(srt.render_srt(thoat), encoding="utf-8")
        fonts_dir = lam / "fonts"
        fonts_dir.mkdir(exist_ok=True)
        shutil.copy2(ITIM_FONT, fonts_dir / ITIM_FONT.name)
        burn = ".khung/phu-de.srt"
    elif phu_de == "karaoke":
        (lam / "phu-de.ass").write_text(karaoke.tao_ass(cac_lich), encoding="utf-8")
        fonts_dir = lam / "fonts"
        fonts_dir.mkdir(exist_ok=True)
        shutil.copy2(ITIM_FONT, fonts_dir / ITIM_FONT.name)
        burn = ".khung/phu-de.ass"
    elif phu_de == "file":
        (thu_muc / "phu-de.srt").write_text(cues, encoding="utf-8")
        files.append("phu-de.srt")
    tam = lam / "video.mp4"
    _chay(lenh_video(danh_sach, tam, fps, burn), run, thu_muc)
    try:
        os.replace(tam, thu_muc / "video.mp4")
    except PermissionError as exc:
        raise media.MediaError("write", "Không ghi được video.mp4 (đang mở ở nơi khác).",
                               "Đóng video.mp4 nếu đang mở trong trình phát rồi chạy lại.") from exc
    return files
