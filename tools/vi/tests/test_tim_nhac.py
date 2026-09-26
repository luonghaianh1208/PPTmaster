"""Công cụ tìm nhạc Openverse: lọc thời lượng, tải bản mp3, ghi nguon.json, một dòng JSON. Không bao giờ gọi mạng thật:
hàm mở địa chỉ (`lay`) luôn là hàm giả."""

import contextlib
import io
import json
import sys
import tempfile
import unittest
import urllib.error
import urllib.parse
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))

import tim_nhac  # noqa: E402

MP3 = b"ID3\x04\x00\x00\x00\x00\x00\x00" + b"\xff\xfb\x90\x00" * 64


def ket_qua(ten, giay, license="by", version="4.0", filetype="mp3", url=None):
    return {"id": ten, "title": f"Bản {ten}", "creator": f"Tác giả {ten}", "license": license, "license_version": version,
            "license_url": f"https://creativecommons.org/licenses/{license}/{version}/",
            "foreign_landing_url": f"https://nguon.example/{ten}", "url": url or f"https://tai.example/{ten}.mp3",
            "duration": None if giay is None else int(giay * 1000), "filetype": filetype}


KET_QUA = [ket_qua("ngan", 30), ket_qua("khong-ro", None), ket_qua("dai-1", 95, license="cc0", version="1.0"),
           ket_qua("dai-2", 61), ket_qua("dai-3", 200)]


class LayGia:
    """Thay urllib.request.urlopen: trả JSON cho địa chỉ API, bytes mp3 cho địa chỉ file; ghi lại địa chỉ đã gọi."""

    def __init__(self, ket=KET_QUA, loi=None, tep=None):
        self.ket, self.loi, self.tep = ket, loi, tep or {}
        self.goi = []

    def __call__(self, req, timeout=None):
        url = req.full_url if hasattr(req, "full_url") else req
        self.goi.append(url)
        if self.loi is not None:
            raise self.loi
        if url.startswith(tim_nhac.API):
            return io.BytesIO(json.dumps({"result_count": len(self.ket), "results": self.ket}).encode("utf-8"))
        return io.BytesIO(self.tep.get(url, MP3))


def chay(argv, lay):
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
        code = tim_nhac.main(argv, lay=lay)
    dong = out.getvalue().splitlines()
    return code, dong


class TimTest(unittest.TestCase):
    def test_goi_openverse_dung_tham_so_va_loc_ban_duoi_60_giay(self):
        lay = LayGia()
        ds = tim_nhac.tim("calm piano", lay=lay)
        self.assertEqual([b["id"] for b in ds], ["dai-1", "dai-2", "dai-3"])
        url = urllib.parse.urlsplit(lay.goi[0])
        self.assertEqual(f"{url.scheme}://{url.netloc}{url.path}", "https://api.openverse.org/v1/audio/")
        q = urllib.parse.parse_qs(url.query)
        self.assertEqual(q, {"q": ["calm piano"], "license": ["cc0,by"], "page_size": ["20"]})

    def test_chi_nhan_ban_mp3(self):
        lay = LayGia([ket_qua("wav", 90, filetype="wav", url="https://tai.example/wav.wav"), ket_qua("mp3", 90)])
        self.assertEqual([b["id"] for b in tim_nhac.tim("x", lay=lay)], ["mp3"])


    def test_chi_tai_qua_https(self):
        lay = LayGia([ket_qua("http", 90, url="http://tai.example/a.mp3"), ket_qua("file", 90, url="file:///C:/x.mp3"),
                      ket_qua("tot", 90)])
        self.assertEqual([b["id"] for b in tim_nhac.tim("x", lay=lay)], ["tot"])


class MainTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.nhac = Path(tmp.name) / "bai" / "nhac"

    def test_tai_ban_dau_ghi_nguon_json_va_in_mot_dong(self):
        lay = LayGia()
        code, dong = chay(["nhạc nền Đà Lạt", "-o", str(self.nhac)], lay)
        self.assertEqual(code, 0, dong)
        self.assertEqual(len(dong), 1)
        kq = json.loads(dong[0])
        self.assertTrue(kq["ready"])
        self.assertIsNone(kq["error"])
        self.assertEqual((self.nhac / "nhac-nen-da-lat.mp3").read_bytes(), MP3)
        self.assertEqual(kq["files"], ["nhac-nen-da-lat.mp3"])
        self.assertIn("https://tai.example/dai-1.mp3", lay.goi)
        nguon = json.loads((self.nhac / "nguon.json").read_text(encoding="utf-8"))
        self.assertEqual(nguon["items"], [{
            "filename": "nhac-nen-da-lat.mp3", "title": "Bản dai-1", "creator": "Tác giả dai-1", "license": "CC0 1.0",
            "license_url": "https://creativecommons.org/licenses/cc0/1.0/", "source_url": "https://nguon.example/dai-1"}])

    def test_so_ban_va_giu_ban_ghi_cu(self):
        self.nhac.mkdir(parents=True)
        (self.nhac / "nguon.json").write_text(json.dumps({"items": [{"filename": "cu.mp3", "title": "Cũ", "creator": "A",
                                                                     "license": "CC0 1.0"}]}), encoding="utf-8")
        code, dong = chay(["calm piano", "-o", str(self.nhac), "--so", "2"], LayGia())
        self.assertEqual(code, 0, dong)
        kq = json.loads(dong[0])
        self.assertEqual(kq["files"], ["calm-piano.mp3", "calm-piano-2.mp3"])
        items = json.loads((self.nhac / "nguon.json").read_text(encoding="utf-8"))["items"]
        self.assertEqual([i["filename"] for i in items], ["cu.mp3", "calm-piano.mp3", "calm-piano-2.mp3"])
        self.assertEqual(items[2]["license"], "CC BY 4.0")

    def test_khong_ghi_de_file_da_co(self):
        self.nhac.mkdir(parents=True)
        (self.nhac / "calm-piano.mp3").write_bytes(b"cu")
        code, dong = chay(["calm piano", "-o", str(self.nhac)], LayGia())
        self.assertEqual(code, 0, dong)
        self.assertEqual(json.loads(dong[0])["files"], ["calm-piano-2.mp3"])
        self.assertEqual((self.nhac / "calm-piano.mp3").read_bytes(), b"cu")

    def test_ban_tai_ve_khong_phai_mp3_thi_bo_qua_sang_ban_sau(self):
        lay = LayGia(tep={"https://tai.example/dai-1.mp3": b"<html>loi</html>"})
        code, dong = chay(["calm piano", "-o", str(self.nhac)], lay)
        self.assertEqual(code, 0, dong)
        kq = json.loads(dong[0])
        items = json.loads((self.nhac / "nguon.json").read_text(encoding="utf-8"))["items"]
        self.assertEqual(items[0]["title"], "Bản dai-2")
        self.assertTrue(kq["warnings"])

    def test_loi_mang_la_mot_dong_json_buoc_mang(self):
        for loi in (urllib.error.URLError("no route"), TimeoutError("het gio"),
                    urllib.error.HTTPError(tim_nhac.API, 503, "busy", {}, None)):
            with self.subTest(loi=type(loi).__name__):
                code, dong = chay(["calm piano", "-o", str(self.nhac)], LayGia(loi=loi))
                self.assertEqual(code, 1)
                self.assertEqual(len(dong), 1)
                kq = json.loads(dong[0])
                self.assertFalse(kq["ready"])
                self.assertEqual(kq["error"]["step"], "mang")
                self.assertFalse((self.nhac / "nguon.json").exists())

    def test_phan_hoi_khong_phai_json_la_loi_mang(self):
        class Hong(LayGia):
            def __call__(self, req, timeout=None):
                return io.BytesIO(b"<html>")
        code, dong = chay(["calm piano", "-o", str(self.nhac)], Hong())
        self.assertEqual(json.loads(dong[0])["error"]["step"], "mang")

    def test_khong_co_ban_phu_hop_la_loi_input(self):
        code, dong = chay(["calm piano", "-o", str(self.nhac)], LayGia([ket_qua("ngan", 20)]))
        self.assertEqual(code, 1)
        self.assertEqual(json.loads(dong[0])["error"]["step"], "input")

    def test_sai_tham_so_la_loi_input(self):
        for argv in ([], ["calm"], ["calm", "-o", str(self.nhac), "--so", "0"], ["   ", "-o", str(self.nhac)],
                     ["!!!", "-o", str(self.nhac)]):
            with self.subTest(argv=argv):
                code, dong = chay(argv, LayGia())
                self.assertEqual(code, 1)
                self.assertEqual(len(dong), 1)
                self.assertEqual(json.loads(dong[0])["error"]["step"], "input")

    def test_khong_ghi_duoc_la_loi_write(self):
        self.nhac.parent.mkdir(parents=True)
        self.nhac.write_bytes(b"day la file, khong phai thu muc")
        code, dong = chay(["calm piano", "-o", str(self.nhac)], LayGia())
        self.assertEqual(code, 1)
        self.assertEqual(json.loads(dong[0])["error"]["step"], "write")

    def test_nguon_json_hong_la_loi_input_khong_ghi_de_khong_len_mang(self):
        self.nhac.mkdir(parents=True)
        (self.nhac / "nguon.json").write_text("{hong", encoding="utf-8")
        lay = LayGia()
        code, dong = chay(["calm piano", "-o", str(self.nhac)], lay)
        self.assertEqual(code, 1)
        self.assertEqual(json.loads(dong[0])["error"]["step"], "input")
        self.assertEqual(lay.goi, [])
        self.assertEqual((self.nhac / "nguon.json").read_text(encoding="utf-8"), "{hong")


class SlugTest(unittest.TestCase):
    def test_bo_dau_tieng_viet(self):
        self.assertEqual(tim_nhac.slug("Nhạc nền Đà Lạt"), "nhac-nen-da-lat")
        self.assertEqual(tim_nhac.slug("  piano, êm   dịu! "), "piano-em-diu")
        self.assertEqual(tim_nhac.slug("!!!"), "")


if __name__ == "__main__":
    unittest.main()
