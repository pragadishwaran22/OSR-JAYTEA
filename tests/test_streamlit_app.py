from __future__ import annotations

import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest


class StreamlitAppTests(unittest.TestCase):
    def test_initial_screen_loads(self):
        app_path = Path(__file__).resolve().parents[1] / "streamlit_app.py"
        app = AppTest.from_file(str(app_path), default_timeout=10).run()
        self.assertFalse(app.exception)
        self.assertEqual("Order planning report", app.title[0].value)
        self.assertEqual("Generate report", app.button[0].label)
        self.assertTrue(app.button[0].disabled)


if __name__ == "__main__":
    unittest.main()
