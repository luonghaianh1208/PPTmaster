"""Chạy bộ test Node của phần logic trong khung chạy thí nghiệm ảo."""

import shutil
import subprocess
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
HAS_NODE = shutil.which("node") is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


@unittest.skipUnless(HAS_NODE, NEED_NODE)
class RuntimeLogicTest(unittest.TestCase):
    def test_runtime_logic_passes_its_node_tests(self):
        proc = subprocess.run([shutil.which("node"), "--test", str(TOOLS_VI / "tests" / "js" / "test_khung.js")],
                              capture_output=True, text=True, encoding="utf-8", errors="replace")
        self.assertEqual(proc.returncode, 0, proc.stdout[-2000:])


if __name__ == "__main__":
    unittest.main()
