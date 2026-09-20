"""Test cho bộ ghép file HTML thí nghiệm ảo."""

import json
import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import PENDULUM  # noqa: E402
from thi_nghiem_parts import build_html, parse, thu_vien  # noqa: E402


class BuildHtmlTest(unittest.TestCase):
    def setUp(self):
        self.model = thu_vien.load("li-con-lac-don", Path("."))
        self.experiment = parse.parse_experiment(PENDULUM, self.model.khai_bao)

    def test_page_is_self_contained_and_offline(self):
        page = build_html.build(self.experiment, self.model)
        for mark in build_html.NETWORK_MARKS + ("<link", "src="):
            self.assertNotIn(mark, page)
        for needle in ("THI_NGHIEM_KHUNG.khoiDong();", "THI_NGHIEM_MO_HINH", 'id="du-lieu"', "<title>Chu kì con lắc đơn</title>"):
            self.assertIn(needle, page)

    def test_embedded_data_round_trips_and_cannot_close_the_script(self):
        self.experiment.meta["tieu-de"] = "H~2~O </script><b>x"
        page = build_html.build(self.experiment, self.model)
        raw = page.split('<script id="du-lieu" type="application/json">', 1)[1].split("</script>", 1)[0]
        self.assertNotIn("<", raw)
        data = json.loads(raw)
        self.assertEqual(data["cauHinh"]["tieuDe"], "H~2~O </script><b>x")
        self.assertEqual(data["khaiBao"]["ma"], "li-con-lac-don")
        self.assertIn("<title>H2O &lt;/script&gt;&lt;b&gt;x</title>", page)

    def test_network_marks_in_model_code_stop_the_build(self):
        self.model.js += "\n// url(x)"
        with self.assertRaises(build_html.BuildError):
            build_html.build(self.experiment, self.model)

if __name__ == "__main__":
    unittest.main()
