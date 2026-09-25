"""Chạy bộ test Node của phần vẽ video giải thích (khung-video.js và tám loại cảnh)."""

import shutil
import subprocess
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
HAS_NODE = shutil.which("node") is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


class SceneRuntimeTest(unittest.TestCase):
    @unittest.skipUnless(HAS_NODE, NEED_NODE)
    def test_runtime_passes_its_node_tests(self):
        proc = subprocess.run([shutil.which("node"), "--test", str(TOOLS_VI / "tests" / "js" / "test_canh.js")],
                              capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
        self.assertEqual(proc.returncode, 0, proc.stdout[-3000:] + proc.stderr[-1500:])


if __name__ == "__main__":
    unittest.main()
