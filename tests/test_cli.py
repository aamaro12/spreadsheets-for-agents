"""The command line prints Markdown or JSON, and exits 2 with a reason when it cannot read."""

from contextlib import redirect_stderr, redirect_stdout
import io
import json
import subprocess
import sys

from openpyxl import Workbook

from sheetscope import __version__
from sheetscope.cli import main
from support import WorkbookCase


class CommandLineTests(WorkbookCase):
    def run_cli(self, *args):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = main([str(arg) for arg in args])
        return code, out.getvalue(), err.getvalue()

    def test_each_step_prints_markdown(self):
        path = self.northwind()
        cases = [
            (["outline", path], "## Model · A1:E12 · 45 cells · 31 formulas"),
            (["find", path, "1,597"], "| Summary!B4 | 1,597 | Revenue | FY2027 | =Model!E4 |"),
            (["find", path, "Growth", "--formulas"], "| Model!C4 | 1,320 | Revenue | FY2025 | =B4*(1+Growth) |"),
            (["read", path, "Model!C4:E4"], "| 4 · Revenue | 1,320 | 1,452 | 1,597 |"),
            (["cell", path, "Model!D4"], "- Reads: Model!C4 (Revenue, FY2025) = 1,320; Growth → Inputs!B4"),
            (["trace", path, "Summary!B4"], "- Customers[Revenue FY2024] → Customers!C2:C6 (5 cells, 0 formulas)"),
            (["trace", path, "Growth", "--dependents", "--depth", "1"], "  - Model!E4 (Revenue, FY2027) = 1,597"),
        ]
        for args, expected in cases:
            with self.subTest(step=args[0]):
                code, out, err = self.run_cli(*args)
                self.assertEqual((code, err), (0, ""))
                self.assertIn(expected, out)

    def test_json_output(self):
        code, out, _ = self.run_cli("cell", self.northwind(), "Model!D4", "--json")
        result = json.loads(out)
        self.assertEqual((code, result["kind"], result["citation"]), (0, "cell", "[northwind.xlsx]Model!D4"))

    def test_problems_exit_2_with_a_reason(self):
        code, out, err = self.run_cli("read", self.northwind(), "Nowhere!A1")
        self.assertEqual((code, out), (2, ""))
        self.assertIn("sheetscope: Nowhere!A1: there is no sheet called Nowhere", err)
        code, _, err = self.run_cli("outline", self.folder / "missing.xlsx")
        self.assertEqual(code, 2)
        self.assertIn("No such workbook", err)

    def test_sheet_names_cannot_break_a_table(self):
        book = Workbook()
        book.active.title = "P|L"
        book.active["A1"] = "Revenue"
        _, out, _ = self.run_cli("find", self.save(book), "Revenue")
        self.assertIn("| 'P\\|L'!A1 | Revenue |", out)

    def test_the_package_runs_as_a_program(self):
        done = subprocess.run([sys.executable, "-m", "sheetscope", "--version"], capture_output=True, text=True)
        self.assertEqual(done.stdout.strip(), f"sheetscope {__version__}")
