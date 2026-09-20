"""Ghép khung chạy, mô hình và cấu hình thành một file HTML tự chứa, không gọi Internet."""

from __future__ import annotations

import html
import json
import re
from pathlib import Path

RUNTIME_DIR = Path(__file__).resolve().parent / "runtime"
FILENAME = "thi-nghiem.html"
NETWORK_MARKS = ("http://", "https://", "//cdn", "@import", "url(")

TEMPLATE = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
{css}
</style>
</head>
<body>
<header class="dau"><h1 id="tieu-de"></h1><p id="phu-de"></p></header>
<main class="bo-cuc">
<section class="mo-phong"><canvas id="khung-ve"></canvas><div id="dieu-khien-chay"></div><div id="so-do"></div></section>
<section class="ben-phai"><div id="tham-so"></div><div id="nhiem-vu"></div></section>
<section class="so-lieu"><div id="bang"></div><canvas id="do-thi"></canvas><div id="khop"></div></section>
</main>
<footer><p id="cong-thuc"></p><p id="dieu-kien"></p><p id="tu-kiem"></p></footer>
<script>
{khung_js}
</script>
<script>
{model_js}
</script>
<script id="du-lieu" type="application/json">{data}</script>
<script>THI_NGHIEM_KHUNG.khoiDong();</script>
</body>
</html>
"""


class BuildError(Exception):
    """File HTML ghép ra không đạt điều kiện chạy không cần mạng."""


def _embed_json(data: dict) -> str:
    """JSON an toàn bên trong thẻ <script>: không để lọt `</script` hay `<!--`."""
    return json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")


def build(experiment, model) -> str:
    khung_js = (RUNTIME_DIR / "khung.js").read_text(encoding="utf-8")
    css = (RUNTIME_DIR / "khung.css").read_text(encoding="utf-8")
    page = TEMPLATE.format(
        title=html.escape(re.sub(r"[~^*]", "", experiment.meta["tieu-de"])),
        css=css,
        khung_js=khung_js,
        model_js=model.js,
        data=_embed_json({"cauHinh": experiment.config(), "khaiBao": model.khai_bao}),
    )
    found = [mark for mark in NETWORK_MARKS if mark in page]
    if found:
        raise BuildError("File HTML chứa dấu hiệu tải từ Internet: " + ", ".join(found))
    return page


def write(experiment, model, folder: Path) -> Path:
    path = folder / FILENAME
    path.write_text(build(experiment, model), encoding="utf-8", newline="\n")
    return path
