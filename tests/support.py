"""Workbooks for the tests, saved with the results Excel would have stored for their formulas."""

from pathlib import Path
import sys
import tempfile
import unittest

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
sys.path.insert(0, str(EXAMPLES))

import northwind  # noqa: E402


class WorkbookCase(unittest.TestCase):
    """A test with its own folder for workbooks."""

    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def save(self, workbook, results=None, name="book.xlsx") -> Path:
        """Save a workbook; results ({sheet: {cell: value}}) become its formulas' saved values."""
        return northwind.save_calculated(workbook, self.folder / name, results or {})

    def northwind(self) -> Path:
        return northwind.build(self.folder / "northwind.xlsx")
