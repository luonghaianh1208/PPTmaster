"""Test dựng trang và chụp khung. Phần Chromium tự bỏ qua nếu máy thiếu Chromium hoặc playwright."""

import importlib.util
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

from video_ma_parts import chup, lich, parse, phong, trang  # noqa: E402

META = "tieu-de: T\nmon: Toán\nlop: 8\n"


def co_chromium() -> bool:
    root = os.environ.get("LOCALAPPDATA")
    if not root or importlib.util.find_spec("playwright") is None:
        return False
    base = Path(root) / "ms-playwright"
    return base.is_dir() and (any(base.glob("chromium-*")) or any(base.glob("chromium_headless_shell-*")))


NEED_CHROMIUM = "máy không có Chromium hoặc playwright"

_CUM_TU = "Nhờ ướt nhẫm quyết định "


def vi_text(n: int) -> str:
    """Chuỗi tiếng Việt có dấu, đủ khoảng trắng, đúng `n` ký tự (không phải một từ dài liền)."""
    s = (_CUM_TU * (n // len(_CUM_TU) + 2))[:n]
    return s[:-1] + "x" if s.endswith(" ") else s


def du_cua(noi_dung: str, giay: float = 6.0, loi: str = "Xin chào các em. Hôm nay học bài mới. Cảm ơn các em."):
    text = f"---\n{META}---\n\n## Cảnh 1\n{noi_dung}loi: {loi}\n"
    canh = parse.parse(text).canh[0]
    giong = lich.GiongInfo(mp3=None, giay=giay, moc_cau=[0.0, 2.0, 4.0], uoc_luong=False, nguon="may")
    plan, _ = lich.dung_lich([canh], [giong])
    return lich.du_lieu_canh(canh, plan[0]), plan[0]


class PageBuildTest(unittest.TestCase):
    def test_page_is_self_contained(self):
        du, _ = du_cua("loai: tieu-de\nchu: Xin chào\n")
        html = trang.dung_trang(du)
        self.assertNotIn("http://", html.replace("http://www.w3.org/2000/svg", ""))
        self.assertNotIn("https://", html)
        self.assertIn("Itim", html)
        self.assertIn("THI_CANH['tieu-de']", html)
        self.assertNotIn("THI_CANH['y-tung-y']", html)
        self.assertIn("khoiDong", html)

    def test_trang_khong_chua_ma_chay_duoc(self):
        du, _ = du_cua('loai: tieu-de\nchu: a < b </script><img src=x onerror=alert(1)> & "c"\n')
        html = trang.dung_trang(du)
        self.assertEqual(html.count("</script>"), html.count("<script>"))
        self.assertNotIn("<img src=x", html)
        self.assertIn("\\u003c/script>", html)

    def test_experiment_page_embeds_khung_and_model(self):
        from thi_nghiem_parts import thu_vien

        model = thu_vien.load("li-con-lac-don", Path("."))
        text = f"---\n{META}---\n\n## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\nloi: Ok.\n"
        canh = parse.parse(text).canh[0]
        giong = lich.GiongInfo(mp3=None, giay=4.0, moc_cau=[0.0], uoc_luong=False, nguon="may")
        plan, _ = lich.dung_lich([canh], [giong])
        html = trang.dung_trang(lich.du_lieu_canh(canh, plan[0], model), model)
        self.assertIn("THI_NGHIEM_KHUNG", html)
        self.assertIn("THI_NGHIEM_MO_HINH", html)
        self.assertLess(html.index("THI_NGHIEM_KHUNG = "), html.index("THI_NGHIEM_MO_HINH = "))


@unittest.skipUnless(co_chromium(), NEED_CHROMIUM)
class ChromiumTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cm = chup.trinh_duyet()
        cls.browser = cls.cm.__enter__()
        cls.page = chup.trang_moi(cls.browser)

    @classmethod
    def tearDownClass(cls):
        cls.cm.__exit__(None, None, None)

    def test_normal_scene_does_not_overflow_and_ink_grows(self):
        du, cl = du_cua("loai: y-tung-y\ntieu-de: Ba bước\ny: Một\ny: Hai\ny: Ba\n")
        html = trang.dung_trang(du)
        self.assertEqual(chup.kiem_tran(self.page, html), [])
        with tempfile.TemporaryDirectory() as tmp:
            n = chup.chup_canh(self.page, html, 6, lich.FPS, Path(tmp), 0)
            self.assertEqual(n, 6)
            frames = sorted(Path(tmp).glob("f*.png"))
            self.assertEqual([f.name for f in frames], [f"f{i:06d}.png" for i in range(6)])
            first = frames[0].read_bytes()
            self.assertEqual(first[:8], b"\x89PNG\r\n\x1a\n")
            final = Path(tmp) / "cuoi.png"
            chup.chup_cuoi(self.page, html, final)
            self.assertGreater(final.stat().st_size, frames[0].stat().st_size)

    def test_itim_renders_every_vietnamese_letter_without_a_fallback_glyph(self):
        html = f"<!doctype html><html><head><style>{phong.font_css()}</style></head><body></body></html>"
        self.page.set_content(html)
        chars = list(phong.CHU_VIET)
        text = "".join(chars)
        loaded = self.page.evaluate(
            "(text) => document.fonts.load(\"40px 'Itim'\", text).then((faces) => faces.length)", text,
        )
        self.assertGreaterEqual(loaded, 1)
        diffs = self.page.evaluate(
            """(chars) => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                return chars.map((c) => {
                    ctx.font = "40px 'Itim', monospace";
                    const w1 = ctx.measureText(c).width;
                    ctx.font = "40px 'Itim', serif";
                    const w2 = ctx.measureText(c).width;
                    return Math.abs(w1 - w2);
                });
            }""",
            chars,
        )
        for d, c in zip(diffs, chars):
            self.assertLess(d, 0.01, c)

    def test_first_frame_of_a_freshly_opened_page_already_has_itim_loaded(self):
        du, _ = du_cua("loai: tieu-de\nchu: " + phong.CHU_VIET[:20] + "\n")
        html = trang.dung_trang(du)
        page = chup.trang_moi(self.browser)
        try:
            chup.mo_trang(page, html)
            chars = list(phong.CHU_VIET)
            diffs = page.evaluate(
                """(chars) => {
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');
                    return chars.map((c) => {
                        ctx.font = "40px 'Itim', monospace";
                        const w1 = ctx.measureText(c).width;
                        ctx.font = "40px 'Itim', serif";
                        const w2 = ctx.measureText(c).width;
                        return Math.abs(w1 - w2);
                    });
                }""",
                chars,
            )
            for d, c in zip(diffs, chars):
                self.assertLess(d, 0.01, c)
        finally:
            page.close()

    def test_scene_page_computed_font_family_starts_with_itim(self):
        du, _ = du_cua("loai: tieu-de\nchu: Xin chào\n")
        html = trang.dung_trang(du)
        chup.mo_trang(self.page, html)
        font_family = self.page.evaluate("() => getComputedStyle(document.querySelector('.chu')).fontFamily")
        self.assertTrue(font_family.startswith("Itim"), font_family)

    def test_overflowing_text_is_reported_by_item_id(self):
        du, _ = du_cua("loai: tieu-de\nchu: " + "A" * 80 + "\n")
        self.assertIn("chu", chup.kiem_tran(self.page, trang.dung_trang(du)))

    def khoang_gach(self, html: str, id_chu: str, y_gach: float) -> float:
        chup.mo_trang(self.page, html)
        self.page.evaluate("window.datThoiDiem(1e6)")
        day = self.page.evaluate(f"document.querySelector('[data-id=\"{id_chu}\"] .trong').getBoundingClientRect().bottom")
        return y_gach - day

    def test_underline_sits_right_under_a_one_line_heading(self):
        for noi_dung, id_chu, y_gach in (("loai: y-tung-y\ntieu-de: Ba bước\ny: Một\n", "tieu-de", 160),
                                         ("loai: khai-niem\nthuat-ngu: Chu kì\ndinh-nghia: Thời gian.\n", "thuat-ngu", 335),
                                         ("loai: tieu-de\nchu: Con lắc đơn\nphu: Vật lí 11\n", "chu", 400)):
            du, _ = du_cua(noi_dung)
            self.assertLess(abs(self.khoang_gach(trang.dung_trang(du), id_chu, y_gach)), 30, id_chu)

    def test_two_line_term_and_heading_fit_their_boxes(self):
        du, _ = du_cua("loai: khai-niem\nthuat-ngu: Chu kì dao động điều hoà của con lắc đơn khi góc lệch nhỏ\n"
                       "dinh-nghia: Thời gian.\n")
        self.assertEqual(chup.kiem_tran(self.page, trang.dung_trang(du)), [])
        du, _ = du_cua("loai: y-tung-y\ntieu-de: Chu kì của con lắc đơn phụ thuộc vào những yếu tố nào và "
                       "không phụ thuộc vào yếu tố nào\ny: Một\n")
        self.assertEqual(chup.kiem_tran(self.page, trang.dung_trang(du)), [])

    def test_list_scenes_fit_their_boxes_at_the_new_limits(self):
        diem = "\n".join(f"diem: {k}, {k * k % 7}" for k in range(12))
        casos = {
            "y-tung-y": "loai: y-tung-y\ntieu-de: Sáu ý\n" + "".join(f"y: {vi_text(60)}\n" for _ in range(6)),
            "cong-thuc": ("loai: cong-thuc\nbieu-thuc: " + vi_text(90) + "\n"
                          + "".join(f"giai-thich: {vi_text(60)}\n" for _ in range(4))),
            "quy-trinh": "loai: quy-trinh\ntieu-de: Năm bước\n" + "".join(f"buoc: {vi_text(50)}\n" for _ in range(5)),
            "so-sanh": ("loai: so-sanh\ntieu-de: So sánh\n" + f"trai: {vi_text(24)}\nphai: {vi_text(24)}\n"
                        + "".join(f"y-trai: {vi_text(60)}\n" for _ in range(4))
                        + "".join(f"y-phai: {vi_text(60)}\n" for _ in range(4))),
            "do-thi": (f"loai: do-thi\ntieu-de: Đồ thị\ntruc-ngang: {vi_text(40)}\ntruc-doc: {vi_text(40)}\n"
                       + diem + "\n"),
            "khai-niem": f"loai: khai-niem\nthuat-ngu: {vi_text(60)}\ndinh-nghia: {vi_text(220)}\n",
            "tieu-de": f"loai: tieu-de\nchu: {vi_text(90)}\nphu: {vi_text(90)}\n",
        }
        for loai, noi_dung in casos.items():
            with self.subTest(loai=loai):
                du, _ = du_cua(noi_dung)
                html = trang.dung_trang(du)
                self.assertEqual(chup.kiem_tran(self.page, html), [])

    def test_hostile_text_renders_as_text(self):
        du, _ = du_cua('loai: tieu-de\nchu: a < b </script><img src=x onerror=alert(1)> & "c"\n')
        html = trang.dung_trang(du)
        chup.mo_trang(self.page, html)
        self.page.evaluate("window.datThoiDiem(1e6)")
        self.assertEqual(self.page.evaluate("document.querySelectorAll('#khung img').length"), 0)
        texto = self.page.evaluate("document.querySelector('[data-id=chu]').textContent")
        self.assertIn("a < b </script><img src=x onerror=alert(1)> & \"c\"", re.sub(r"\s+", " ", texto))

    def test_frames_are_repeatable(self):
        du, _ = du_cua("loai: khai-niem\nthuat-ngu: Chu kì\ndinh-nghia: Thời gian một dao động.\n")
        html = trang.dung_trang(du)
        with tempfile.TemporaryDirectory() as tmp:
            chup.chup_canh(self.page, html, 4, lich.FPS, Path(tmp), 0)
            chup.chup_canh(self.page, html, 4, lich.FPS, Path(tmp), 10)
            for i in range(4):
                self.assertEqual((Path(tmp) / f"f{i:06d}.png").read_bytes(), (Path(tmp) / f"f{10 + i:06d}.png").read_bytes())

    def test_experiment_scene_draws_the_model(self):
        from thi_nghiem_parts import thu_vien

        model = thu_vien.load("li-con-lac-don", Path("."))
        text = (f"---\n{META}---\n\n## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 0 chieu-dai 0.4\n"
                "tham-so: 5 chieu-dai 1.6\ndo: chu-ki\nloi: Ok.\n")
        canh = parse.parse(text).canh[0]
        giong = lich.GiongInfo(mp3=None, giay=6.0, moc_cau=[0.0], uoc_luong=False, nguon="may")
        plan, _ = lich.dung_lich([canh], [giong])
        html = trang.dung_trang(lich.du_lieu_canh(canh, plan[0], model), model)
        self.assertEqual(chup.kiem_tran(self.page, html), [])
        self.page.evaluate("window.datThoiDiem(3)")
        dong = self.page.evaluate("document.querySelector('[data-id=do-0]').textContent")
        self.assertRegex(dong, r"Chu kì T = \d,\d{3} s")
        self.page.evaluate("window.datThoiDiem(1e6)")
        cuoi = self.page.evaluate("document.querySelector('[data-id=do-0]').textContent")
        self.assertIn("2,539", cuoi)

    def test_preview_shows_the_scene_at_its_final_parameter_value(self):
        from thi_nghiem_parts import thu_vien

        model = thu_vien.load("li-con-lac-don", Path("."))
        text = (f"---\n{META}---\n\n## Cảnh 1\nloai: thi-nghiem\nmau: li-con-lac-don\ntham-so: 0 chieu-dai 0.4\n"
                "tham-so: 5 chieu-dai 1.6\ndo: chu-ki\nloi: Ok.\n")
        canh = parse.parse(text).canh[0]
        giong = lich.GiongInfo(mp3=None, giay=6.0, moc_cau=[0.0], uoc_luong=False, nguon="may")
        plan, _ = lich.dung_lich([canh], [giong])
        html = trang.dung_trang(lich.du_lieu_canh(canh, plan[0], model), model)
        chup.mo_trang(self.page, html)
        self.page.evaluate("window.datThoiDiem(window.THI_VIDEO.thoiDiemCuoi())")
        ts0 = self.page.evaluate("document.querySelector('[data-id=ts-0]').textContent")
        self.assertIn("1,60 m", ts0)


def png(rong: int, cao: int) -> bytes:
    """PNG sọc ngang tự dựng bằng thư viện chuẩn (để thấy ảnh chuyển động)."""
    import struct
    import zlib

    hang = [b"\x00" + (bytes((200, 90, 60)) if (y // 40) % 2 else bytes((60, 120, 200))) * rong for y in range(cao)]
    def khoi(loai, du):
        return struct.pack(">I", len(du)) + loai + du + struct.pack(">I", zlib.crc32(loai + du) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + khoi(b"IHDR", struct.pack(">IIBBBBB", rong, cao, 8, 2, 0, 0, 0))
            + khoi(b"IDAT", zlib.compress(b"".join(hang), 9)) + khoi(b"IEND", b""))


def anh_gia(rong: int, cao: int) -> dict:
    import base64
    return {"dataUrl": "data:image/png;base64," + base64.b64encode(png(rong, cao)).decode("ascii"),
            "nguon": "Ảnh: Tác giả thử · CC BY 4.0 · Wikimedia", "rong": rong, "cao": cao}


def du_hinh(noi_dung: str, tai_nguyen: dict, so: int = 1, giay: float = 6.0,
            loi: str = "Xin chào các em. Hôm nay học bài mới. Cảm ơn các em."):
    truoc = "".join(f"## Cảnh {k}\nloai: tieu-de\nchu: Mở đầu\nloi: Chào.\n\n" for k in range(1, so))
    text = f"---\n{META}---\n\n{truoc}## Cảnh {so}\n{noi_dung}loi: {loi}\n"
    canh = parse.parse(text).canh[so - 1]
    giong = lich.GiongInfo(mp3=None, giay=giay, moc_cau=[0.0, 2.0, 4.0], uoc_luong=False, nguon="may")
    plan, _ = lich.dung_lich([canh], [giong])
    return lich.du_lieu_canh(canh, plan[0], None, {**tai_nguyen, "meta": {}})


@unittest.skipUnless(co_chromium(), NEED_CHROMIUM)
class PictureMotionChromiumTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from video_ma_parts import hinh
        cls.hinh = hinh.doc("flask")
        cls.cm = chup.trinh_duyet()
        cls.browser = cls.cm.__enter__()
        cls.page = chup.trang_moi(cls.browser)

    @classmethod
    def tearDownClass(cls):
        cls.cm.__exit__(None, None, None)

    def muc(self, id_muc: str) -> dict:
        return self.page.evaluate(
            "(id) => THI_CANH[DU_CANH.loai].muc(DU_CANH).filter((m) => m.id === id)[0]", id_muc)

    def test_picture_column_scenes_fit_at_the_text_limits(self):
        casos = {
            "khai-niem": f"loai: khai-niem\nthuat-ngu: {vi_text(60)}\ndinh-nghia: {vi_text(220)}\n",
            "y-tung-y": f"loai: y-tung-y\ntieu-de: {vi_text(90)}\n" + "".join(f"y: {vi_text(60)}\n" for _ in range(6)),
            "cong-thuc": ("loai: cong-thuc\nbieu-thuc: " + vi_text(90) + "\n"
                          + "".join(f"giai-thich: {vi_text(60)}\n" for _ in range(4))),
            "tieu-de": f"loai: tieu-de\nchu: {vi_text(90)}\nphu: {vi_text(90)}\n",
        }
        for loai, noi_dung in casos.items():
            for kieu, tai in (("hinh", {"hinh": self.hinh}), ("anh", {"anh": anh_gia(600, 1200)})):
                with self.subTest(loai=loai, kieu=kieu):
                    du = du_hinh(noi_dung + ("hinh: flask\n" if kieu == "hinh" else "anh: a.png\n"), tai)
                    self.assertEqual(chup.kiem_tran(self.page, trang.dung_trang(du)), [])

    def test_minh_hoa_three_pictures_with_long_labels_fit(self):
        nhan = vi_text(30)
        noi_dung = "loai: minh-hoa\ntieu-de: " + vi_text(90) + "\n" + "".join(f"hinh: flask | {nhan}\n" for _ in range(3))
        du = du_hinh(noi_dung, {"hinhs": [{**self.hinh, "nhan": nhan} for _ in range(3)]})
        self.assertEqual(chup.kiem_tran(self.page, trang.dung_trang(du)), [])

    def test_photo_scene_keeps_the_photo_ratio_inside_its_area(self):
        for rong, cao in ((600, 1200), (2000, 800)):
            with self.subTest(rong=rong, cao=cao):
                du = du_hinh(f"loai: anh\nanh: a.png\nchu-thich: {vi_text(90)}\nnguon: Tôi\n", {"anh": anh_gia(rong, cao)})
                html = trang.dung_trang(du)
                self.assertEqual(chup.kiem_tran(self.page, html), [])
                m = self.muc("anh")
                self.page.evaluate("(t) => window.datThoiDiem(t)", m["batDau"])
                r = self.page.evaluate("""() => { const i = document.querySelector('.anh img');
                    const b = i.getBoundingClientRect();
                    return {x: b.left, y: b.top, w: b.width, h: b.height, nw: i.naturalWidth, nh: i.naturalHeight}; }""")
                self.assertEqual((r["nw"], r["nh"]), (rong, cao))
                self.assertGreaterEqual(r["x"], 80 - 0.5)
                self.assertGreaterEqual(r["y"], 70 - 0.5)
                self.assertLessEqual(r["x"] + r["w"], 1200 + 0.5)
                self.assertLessEqual(r["y"] + r["h"], 560 + 0.5)
                self.assertAlmostEqual(r["w"] / r["h"], rong / cao, delta=0.01 * rong / cao)
                nguon = self.page.evaluate("document.querySelector('.anh .nguon').textContent")
                self.assertIn("Tác giả thử", nguon)
                cuoi = self.page.screenshot(type="png")
                self.page.evaluate("(t) => window.datThoiDiem(t)", du["thoiLuong"] - 0.2)
                self.assertNotEqual(cuoi, self.page.screenshot(type="png"), "ảnh phải chuyển động (Ken Burns)")

    def test_icon_is_drawn_stroke_by_stroke(self):
        du = du_hinh("loai: y-tung-y\ntieu-de: Dụng cụ\nhinh: flask\ny: Bình tam giác\n", {"hinh": self.hinh})
        chup.mo_trang(self.page, trang.dung_trang(du))
        m = self.muc("hinh")
        self.assertGreaterEqual(m["thoiLuong"], 1.2)
        self.page.evaluate("(t) => window.datThoiDiem(t)", m["batDau"] + 0.5 * m["thoiLuong"])
        giua = self.page.screenshot(type="png")
        self.page.evaluate("(t) => window.datThoiDiem(t)", du["thoiLuong"] - 0.2)
        self.assertNotEqual(giua, self.page.screenshot(type="png"))
        mo = self.page.evaluate("""() => Array.from(document.querySelectorAll('g.hinh > *'))
            .map((e) => getComputedStyle(e).opacity)""")
        self.assertEqual(set(mo), {"1"})

    def test_pen_hand_sits_on_the_pen_tip_while_writing(self):
        noi_dung = "loai: y-tung-y\ntieu-de: Ba bước\ny: Nhờ ướt nhẫm quyết định mọi thứ\ny: Hai bước nhỏ\n"
        ngoi_chu = """(id) => {
            const tay = document.getElementById('ban-tay');
            const s = document.querySelector('[data-id="' + id + '"] .ngoi').getBoundingClientRect();
            const n = document.getElementById('ngoi-but').getBoundingClientRect();
            return {hien: getComputedStyle(tay).display !== 'none',
                    d: Math.hypot((n.left + n.width / 2) - s.left, (n.top + n.height / 2) - (s.top + 0.8 * s.height))};
        }"""
        for tai, cac_muc in (({}, ("tieu-de", "y-0", "y-1")), ({"hinh": self.hinh}, ("tieu-de", "y-1"))):
            du = du_hinh(noi_dung + ("hinh: flask\n" if tai else ""), tai)
            chup.mo_trang(self.page, trang.dung_trang(du))
            for id_muc in cac_muc:
                with self.subTest(hinh=bool(tai), muc=id_muc):
                    m = self.muc(id_muc)
                    self.page.evaluate("(t) => window.datThoiDiem(t)", m["batDau"] + 0.5 * m["thoiLuong"])
                    kq = self.page.evaluate(ngoi_chu, id_muc)
                    self.assertTrue(kq["hien"])
                    self.assertLess(kq["d"], 40)
            self.page.evaluate("(t) => window.datThoiDiem(t)", du["thoiLuong"] - 1 / 30)
            self.assertEqual(self.page.evaluate("getComputedStyle(document.getElementById('ban-tay')).display"), "none")
        m = self.muc("hinh")
        self.page.evaluate("(t) => window.datThoiDiem(t)", m["batDau"] + 0.5 * m["thoiLuong"])
        trong = self.page.evaluate("""() => {
            const g = document.querySelector('g.hinh').getBoundingClientRect();
            const n = document.getElementById('ngoi-but').getBoundingClientRect();
            const x = n.left + n.width / 2, y = n.top + n.height / 2;
            return x >= g.left - 10 && x <= g.right + 10 && y >= g.top - 10 && y <= g.bottom + 10; }""")
        self.assertTrue(trong, "giữa lúc vẽ hình, ngòi bút phải nằm trên hình")

    def test_board_wipe_uncovers_the_previous_scene(self):
        import base64
        du = du_hinh("loai: khai-niem\nthuat-ngu: Chu kì\ndinh-nghia: Thời gian.\n", {}, so=2)
        self.assertTrue(du["co"]["lauBang"])
        du["nenTruoc"] = "data:image/png;base64," + base64.b64encode(png(1280, 720)).decode("ascii")
        chup.mo_trang(self.page, trang.dung_trang(du))
        thay = """() => { const n = document.getElementById('nen-truoc');
            if (!n || getComputedStyle(n).display === 'none') { return 0; }
            const m = /inset\\(0(?:px)? 0(?:px)? 0(?:px)? ([\\d.]+)px\\)/.exec(n.style.clipPath);
            return 1280 - (m ? Number(m[1]) : 0); }"""
        self.page.evaluate("window.datThoiDiem(0.05)")
        self.assertGreater(self.page.evaluate(thay), 1000)
        self.assertEqual(self.page.evaluate("document.getElementById('ban-tay').getAttribute('data-kieu')"), "gie")
        self.page.evaluate("window.datThoiDiem(0.6)")
        self.assertEqual(self.page.evaluate(thay), 0)

    def test_camera_zooms_during_the_scene_and_is_home_on_the_last_frame(self):
        du = du_hinh("loai: minh-hoa\ntieu-de: Dụng cụ\nhinh: flask | Bình\nhinh: flask | Bình\nhinh: flask | Bình\n",
                     {"hinhs": [{**self.hinh, "nhan": "Bình"} for _ in range(3)]})
        html = trang.dung_trang(du)
        chup.mo_trang(self.page, html)
        m = self.muc("hinh-1")
        tf = "getComputedStyle(document.getElementById('bang')).transform"
        self.page.evaluate("(t) => window.datThoiDiem(t)", m["batDau"] + 0.8 * m["thoiLuong"])
        giua = self.page.evaluate(tf)
        self.assertNotEqual(giua, "none")
        self.assertGreater(float(giua[len("matrix("):].split(",")[0]), 1.05)
        self.page.evaluate("(t) => window.datThoiDiem(t)", du["thoiLuong"] - 1 / 30)
        self.assertIn(self.page.evaluate(tf), ("none", "matrix(1, 0, 0, 1, 0, 0)"))
        t = m["batDau"] + 0.5 * m["thoiLuong"]
        self.page.evaluate("(t) => window.datThoiDiem(t)", t)
        mot = self.page.screenshot(type="png")
        self.page.evaluate("(t) => window.datThoiDiem(t)", 0.1)
        self.page.evaluate("(t) => window.datThoiDiem(t)", t)
        self.assertEqual(mot, self.page.screenshot(type="png"), "cùng t phải cho cùng khung")

    def test_kiem_tran_still_catches_overflow_with_a_picture(self):
        du = du_hinh("loai: y-tung-y\ntieu-de: T\ny: " + "A" * 60 + "\nhinh: flask\n", {"hinh": self.hinh})
        self.assertIn("y-0", chup.kiem_tran(self.page, trang.dung_trang(du)))


class ChromiumMissingTest(unittest.TestCase):
    def test_missing_playwright_is_a_chromium_error(self):
        import builtins
        real_import = builtins.__import__

        def no_pw(name, *args, **kwargs):
            if name.startswith("playwright"):
                raise ImportError("no playwright")
            return real_import(name, *args, **kwargs)

        builtins.__import__ = no_pw
        try:
            with self.assertRaises(chup.MediaError) as caught:
                with chup.trinh_duyet():
                    pass
        finally:
            builtins.__import__ = real_import
        self.assertEqual(caught.exception.step, "chromium")
        self.assertIn("pptmaster.ps1", caught.exception.fix)


if __name__ == "__main__":
    unittest.main()
