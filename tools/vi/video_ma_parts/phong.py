"""Font Itim đóng gói cho khung hình và phụ đề: bộ đọc bảng cmap và khối @font-face nhúng data:."""

from __future__ import annotations

import base64
import struct
from pathlib import Path

FONT = Path(__file__).resolve().parent / "runtime" / "fonts" / "Itim-Regular.ttf"
TEN = "Itim"

_THUONG = "ạảãàáâậầấẩẫăặằắẳẵẹẻẽèéêệềếểễịỉĩìíọỏõòóôộồốổỗơợờớởỡụủũùúưựừứửữỵỷỹỳýđ"
CHU_VIET = _THUONG + _THUONG.upper()


def _doan_dinh_dang_4(data: bytes, offset: int) -> set[int]:
    seg_count_x2 = struct.unpack_from(">H", data, offset + 6)[0]
    seg_count = seg_count_x2 // 2
    end_codes_off = offset + 14
    end_codes = struct.unpack_from(f">{seg_count}H", data, end_codes_off)
    start_codes_off = end_codes_off + seg_count_x2 + 2
    start_codes = struct.unpack_from(f">{seg_count}H", data, start_codes_off)
    id_delta_off = start_codes_off + seg_count_x2
    id_deltas = struct.unpack_from(f">{seg_count}h", data, id_delta_off)
    id_range_off_off = id_delta_off + seg_count_x2
    id_range_offsets = struct.unpack_from(f">{seg_count}H", data, id_range_off_off)

    ma = set()
    for i in range(seg_count):
        start, end = start_codes[i], end_codes[i]
        if start == 0xFFFF and end == 0xFFFF:
            continue
        id_range_offset = id_range_offsets[i]
        for c in range(start, end + 1):
            if id_range_offset == 0:
                glyph_id = (c + id_deltas[i]) & 0xFFFF
            else:
                addr = id_range_off_off + i * 2 + id_range_offset + (c - start) * 2
                if addr + 2 > len(data):
                    continue
                glyph_id = struct.unpack_from(">H", data, addr)[0]
                if glyph_id != 0:
                    glyph_id = (glyph_id + id_deltas[i]) & 0xFFFF
            if glyph_id != 0:
                ma.add(c)
    return ma


def _doan_dinh_dang_12(data: bytes, offset: int) -> set[int]:
    so_nhom = struct.unpack_from(">I", data, offset + 12)[0]
    ma = set()
    pos = offset + 16
    for _ in range(so_nhom):
        start, end, _glyph_dau = struct.unpack_from(">III", data, pos)
        ma.update(range(start, end + 1))
        pos += 12
    return ma


def bang_ma(path) -> set:
    """Đọc bảng `cmap` (định dạng 4 và 12) của một font TrueType, trả về tập mã Unicode có mặt."""
    data = Path(path).read_bytes()
    so_bang = struct.unpack_from(">H", data, 4)[0]
    cmap_offset = None
    pos = 12
    for _ in range(so_bang):
        tag, _checksum, offset, _length = struct.unpack_from(">4sIII", data, pos)
        if tag == b"cmap":
            cmap_offset = offset
        pos += 16
    if cmap_offset is None:
        return set()

    so_bang_con = struct.unpack_from(">H", data, cmap_offset + 2)[0]
    ma = set()
    pos = cmap_offset + 4
    for _ in range(so_bang_con):
        _platform, _encoding, sub_offset = struct.unpack_from(">HHI", data, pos)
        sub_offset += cmap_offset
        dinh_dang = struct.unpack_from(">H", data, sub_offset)[0]
        if dinh_dang == 4:
            ma |= _doan_dinh_dang_4(data, sub_offset)
        elif dinh_dang == 12:
            ma |= _doan_dinh_dang_12(data, sub_offset)
        pos += 8
    return ma


def font_css() -> str:
    b64 = base64.b64encode(FONT.read_bytes()).decode("ascii")
    return f"@font-face{{font-family:'{TEN}';src:url(data:font/ttf;base64,{b64});}}"
