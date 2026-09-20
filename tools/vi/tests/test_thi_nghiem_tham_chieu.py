"""Test định luật cho bản tính lại bằng Python của các mô hình thí nghiệm ảo."""

import math
import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import tham_chieu  # noqa: E402

PROJECTILE = {"van-toc-dau": 20, "goc": 45, "do-cao-dau": 0, "g": 9.8}
PENDULUM_DEFAULTS = {"chieu-dai": 1.0, "g": 9.8, "goc-lech": 8, "khoi-luong": 0.2}
CIRCUIT = {"suat-dien-dong": 12, "dien-tro-1": 10, "dien-tro-2": 20, "kieu-mac": "noi-tiep"}
TITRATION = {"loai-acid": "hcl", "nong-do-acid": 0.1, "the-tich-acid": 20, "nong-do-base": 0.1,
             "the-tich-base": 0, "chi-thi": "phenolphtalein"}
RATE = {"nong-do": 0.1, "nhiet-do": 25, "xuc-tac": "khong"}


class ScienceLawTest(unittest.TestCase):
    """Định luật trên bản Python tham chiếu; test đối chiếu JavaScript bảo đảm JS khớp bản này."""

    def test_projectile_conserves_energy_and_peaks_at_45_degrees(self):
        p = {**PROJECTILE, "van-toc-dau": 20, "goc": 60, "do-cao-dau": 5}
        d = tham_chieu.li_nem_xien(p)
        vy = p["van-toc-dau"] * math.sin(math.radians(60))
        self.assertAlmostEqual(p["g"] * d["do-cao-cuc-dai"], p["g"] * 5 + vy * vy / 2, places=9)
        ranges = {goc: tham_chieu.li_nem_xien({**PROJECTILE, "goc": goc})["tam-xa"] for goc in (30, 45, 60)}
        self.assertAlmostEqual(ranges[30], ranges[60], places=9)
        self.assertGreater(ranges[45], ranges[30])

    def test_pendulum_period_squared_is_proportional_to_length_and_ignores_mass(self):
        base = PENDULUM_DEFAULTS
        short = tham_chieu.li_con_lac_don({**base, "chieu-dai": 0.5})["chu-ki"]
        long = tham_chieu.li_con_lac_don({**base, "chieu-dai": 2.0})["chu-ki"]
        self.assertAlmostEqual(long / short, 2.0, places=12)
        heavy = tham_chieu.li_con_lac_don({**base, "khoi-luong": 1.0})["chu-ki"]
        self.assertEqual(heavy, tham_chieu.li_con_lac_don(base)["chu-ki"])

    def test_circuit_obeys_kirchhoff(self):
        base = CIRCUIT
        series = tham_chieu.li_mach_ohm(base)
        self.assertAlmostEqual(series["hieu-dien-the-1"] + series["hieu-dien-the-2"], base["suat-dien-dong"], places=12)
        parallel = tham_chieu.li_mach_ohm({**base, "kieu-mac": "song-song"})
        self.assertAlmostEqual(parallel["cuong-do-1"] + parallel["cuong-do-2"], parallel["cuong-do-mach-chinh"], places=12)
        self.assertLess(parallel["dien-tro-tuong-duong"], min(base["dien-tro-1"], base["dien-tro-2"]))

    def test_titration_has_textbook_landmarks(self):
        base = TITRATION
        ph = lambda **change: tham_chieu.hoa_chuan_do({**base, **change})["ph"]  # noqa: E731
        self.assertAlmostEqual(ph(**{"the-tich-base": 0}), 1.0, places=6)
        self.assertAlmostEqual(ph(**{"the-tich-base": 20}), 7.0, places=6)
        weak = {"loai-acid": "ch3cooh"}
        self.assertAlmostEqual(ph(**weak, **{"the-tich-base": 10}), -math.log10(1.75e-5), places=2)
        self.assertGreater(ph(**weak, **{"the-tich-base": 20}), 8.0)
        curve = [ph(**weak, **{"the-tich-base": volume / 2}) for volume in range(0, 101)]
        self.assertEqual(curve, sorted(curve))

    def test_no2_equilibrium_follows_le_chatelier(self):
        value = lambda t, p: tham_chieu.hoa_can_bang_no2({"nhiet-do": t, "ap-suat": p})  # noqa: E731
        self.assertGreater(value(60, 1)["phan-mol-no2"], value(20, 1)["phan-mol-no2"])
        self.assertLess(value(25, 4)["phan-mol-no2"], value(25, 1)["phan-mol-no2"])
        self.assertGreater(value(25, 4)["nong-do-no2"], value(25, 1)["nong-do-no2"])
        here = value(25, 1)
        x = here["phan-mol-no2"]
        self.assertAlmostEqual(x * x * 1 / (1 - x), here["kp"], places=12)
        self.assertAlmostEqual(here["kp"], 0.146, places=2)

    def test_rate_follows_concentration_arrhenius_and_catalyst(self):
        base = RATE
        time = lambda **change: tham_chieu.hoa_toc_do({**base, **change})["thoi-gian"]  # noqa: E731
        self.assertAlmostEqual(time(), 40.0, places=9)
        self.assertAlmostEqual(time(**{"nong-do": 0.2}), 20.0, places=9)
        ratio = time() / time(**{"nhiet-do": 35})
        self.assertAlmostEqual(ratio, math.exp(50000 / 8.314462618 * (1 / 298.15 - 1 / 308.15)), places=9)
        self.assertLess(time(**{"xuc-tac": "co"}), time())

    def test_function_extrema_and_tangent(self):
        d = tham_chieu.toan_ham_so({"a": 0, "b": 2, "c": -8, "d": 1, "x0": 2})
        self.assertEqual((d["hoanh-do-cuc-tieu"], d["hoanh-do-cuc-dai"], d["he-so-goc"]), (2.0, None, 0.0))
        none = tham_chieu.toan_ham_so({"a": 1, "b": 0, "c": 3, "d": 0, "x0": 0})
        self.assertEqual((none["hoanh-do-cuc-dai"], none["hoanh-do-cuc-tieu"]), (None, None))

    def test_frequency_approaches_probability(self):
        for phep_thu, xac_suat in (("dong-xu-ngua", 1 / 2), ("hai-xuc-xac-tong-7", 1 / 6)):
            d = tham_chieu.toan_xac_suat({"phep-thu": phep_thu, "so-lan": 100000, "hat-giong": 3})
            self.assertLess(abs(d["tan-suat"] - xac_suat), 0.01)
            self.assertEqual(d["tan-so"], round(d["tan-suat"] * 100000))

if __name__ == "__main__":
    unittest.main()
