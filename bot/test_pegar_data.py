import tempfile
import unittest
import sys
import types
from pathlib import Path
from unittest.mock import patch

playwright_stub = types.ModuleType("playwright")
playwright_sync_api_stub = types.ModuleType("playwright.sync_api")


class PlaywrightTimeoutError(Exception):
    pass


playwright_sync_api_stub.TimeoutError = PlaywrightTimeoutError
playwright_sync_api_stub.sync_playwright = lambda: None
sys.modules.setdefault("playwright", playwright_stub)
sys.modules.setdefault("playwright.sync_api", playwright_sync_api_stub)

import pegar_data


class FailingScreenshotPage:
    def screenshot(self, **kwargs):
        raise PlaywrightTimeoutError("screenshot timed out")


class PegarDataTests(unittest.TestCase):
    def test_capture_timeout_does_not_stop_the_workflow(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(pegar_data, "CAPTURES_DIR", Path(directory)):
                pegar_data.capture(FailingScreenshotPage(), "04_rango_fechas.png")

            self.assertFalse((Path(directory) / "04_rango_fechas.png").exists())


if __name__ == "__main__":
    unittest.main()
