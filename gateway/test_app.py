from __future__ import annotations

import unittest

from gateway import app as gateway_app


class GatewayRootTests(unittest.TestCase):
    def test_weekly_root_uses_current_suite_project(self) -> None:
        self.assertEqual(gateway_app.WEEKLY_ROOT.name, "new-product-monitor")
        self.assertTrue(gateway_app.WEEKLY_ROOT.is_dir())


if __name__ == "__main__":
    unittest.main()
