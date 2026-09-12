from __future__ import annotations

import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest


class StreamlitAppTests(unittest.TestCase):
    def test_initial_screen_loads(self):
        app_path = Path(__file__).resolve().parents[1] / "streamlit_app.py"
        app = AppTest.from_file(str(app_path), default_timeout=10).run()
        self.assertFalse(app.exception)
        self.assertEqual("Order Planning Studio", app.title[0].value)
        self.assertEqual("Generate report", app.button[0].label)
        self.assertTrue(app.button[0].disabled)
        self.assertTrue(any("JAY TEA" in item.value for item in app.markdown))
        self.assertTrue(any("JAY TEA** Order intelligence" in item.value for item in app.markdown))
        self.assertTrue(
            any("Version 0.2" in item.value and "Local" not in item.value for item in app.markdown)
        )


if __name__ == "__main__":
    unittest.main()
