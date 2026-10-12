"""The example workbook in the repository is the one examples/northwind.py builds."""

from sheetscope import load
from support import EXAMPLES, WorkbookCase


class ExampleTests(WorkbookCase):
    def test_the_saved_example_matches_a_fresh_build(self):
        saved, fresh = load(EXAMPLES / "northwind.xlsx"), load(self.northwind())
        self.assertEqual(saved.outline()["warnings"], [])
        for name in fresh.sheets:
            with self.subTest(sheet=name):
                self.assertEqual(saved.read(name)["cells"], fresh.read(name)["cells"])
        self.assertEqual(saved.outline()["names"], fresh.outline()["names"])
        self.assertEqual(saved.outline()["tables"], fresh.outline()["tables"])
