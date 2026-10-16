"""References are parsed, quoted and cited the way Excel writes them."""

import unittest

from sheetscope.refs import MAX_COL, MAX_ROW, Area, cite, excel_formula, formula_refs, parse_area, quote, split_ref


class AreaTests(unittest.TestCase):
    def test_a1_references(self):
        self.assertEqual(parse_area("B3"), Area(3, 2, 3, 2))
        self.assertEqual(parse_area("$B$3:$F$10"), Area(3, 2, 10, 6))
        self.assertEqual(parse_area("F10:b3"), Area(3, 2, 10, 6))
        self.assertEqual(parse_area("A:C"), Area(1, 1, MAX_ROW, 3))
        self.assertEqual(parse_area("$3:5"), Area(3, 1, 5, MAX_COL))
        for text in ["Growth", "B0", "XFE1", "A1:B", "", "R1C1", "A1048577"]:
            with self.subTest(text=text):
                self.assertIsNone(parse_area(text))

    def test_areas_print_as_excel_does(self):
        for text in ["B3", "B3:F10", "A:C", "3:5"]:
            with self.subTest(text=text):
                self.assertEqual(parse_area(text).a1(), text)
        self.assertTrue(Area(3, 2, 10, 6).covers(4, 6))
        self.assertFalse(Area(3, 2, 10, 6).covers(11, 2))
        self.assertEqual(Area(1, 1, 10, 10).clip(Area(5, 5, 20, 20)), Area(5, 5, 10, 10))
        self.assertIsNone(Area(1, 1, 2, 2).clip(Area(5, 5, 6, 6)))


class SheetNameTests(unittest.TestCase):
    def test_quoting(self):
        cases = {"Model": "Model", "P&L 2025": "'P&L 2025'", "Bob's": "'Bob''s'", "2025": "'2025'", "A1": "'A1'"}
        for name, quoted in cases.items():
            with self.subTest(name=name):
                self.assertEqual(quote(name), quoted)

    def test_splitting(self):
        cases = {
            "'P&L 2025'!A1": (None, "P&L 2025", "A1"),
            "Model!B3:C4": (None, "Model", "B3:C4"),
            "A1": (None, None, "A1"),
            "'Bob''s'!A1": (None, "Bob's", "A1"),
            "Data:Other!C3": (None, "Data:Other", "C3"),
            "[1]Ext!A1": ("1", "Ext", "A1"),
            "'[Book 1.xlsx]My Sheet'!A1": ("Book 1.xlsx", "My Sheet", "A1"),
            "[northwind.xlsx]Model!D4": ("northwind.xlsx", "Model", "D4"),
            "'[Model [v2] final.xlsx]F'!B1": ("Model [v2] final.xlsx", "F", "B1"),
            "[Book 1.xlsx]'My Sheet'!A1": ("Book 1.xlsx", "My Sheet", "A1"),  # as agents often write citations
            "[book.xlsx]'Bob''s'!A1": ("book.xlsx", "Bob's", "A1"),
            "Customers[Revenue]": (None, None, "Customers[Revenue]"),
        }
        for text, parts in cases.items():
            with self.subTest(text=text):
                self.assertEqual(split_ref(text), parts)

    def test_citations_use_excels_external_reference_form(self):
        self.assertEqual(cite("northwind.xlsx", "Model", "D4"), "[northwind.xlsx]Model!D4")
        self.assertEqual(cite("my book.xlsx", "P&L", "A1"), "'[my book.xlsx]P&L'!A1")
        self.assertEqual(split_ref(cite("my book.xlsx", "Bob's", "A1")), ("my book.xlsx", "Bob's", "A1"))


class FormulaTests(unittest.TestCase):
    def test_references_in_a_formula(self):
        refs, warning = formula_refs("=SUM('P&L 2025'!A1:B2, Data:Other!C3)+Growth*Customers[Revenue]")
        self.assertEqual(refs, ["'P&L 2025'!A1:B2", "Data:Other!C3", "Growth", "Customers[Revenue]"])
        self.assertIsNone(warning)

    def test_formulas_whose_references_are_only_known_after_calculation(self):
        refs, warning = formula_refs('=INDIRECT("A1")+B2')
        self.assertEqual(refs, ["B2"])
        self.assertIn("INDIRECT", warning)
        for formula in ("=SUM(Model!A1#)", "=SUM(_xlfn.ANCHORARRAY(Model!A1))"):
            with self.subTest(formula=formula):
                refs, warning = formula_refs(formula)
                self.assertEqual(refs, ["Model!A1"])
                self.assertIn("spill", warning)
        self.assertEqual(formula_refs('="unterminated')[0], [])
        self.assertIn("could not be parsed", formula_refs('="unterminated')[1])

    def test_malformed_formulas_strings_and_lambda_parameters_do_not_mislead(self):
        self.assertEqual(formula_refs("=SUM(B1:B2))"), ([], "could not be parsed"))
        self.assertEqual(formula_refs('="see B2#"&A1'), (["A1"], None))
        self.assertEqual(formula_refs("=_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)"), (["A1"], None))
        self.assertEqual(formula_refs("=Model!#REF!+A1"), (["A1"], "reads a deleted range (#REF!)"))

    def test_formulas_that_differ_only_in_cells_read_the_same_way(self):
        tricky = [
            "=LOG10(A1)+Rate_2024*FY2024_Plan",
            "=SUM('Q1 2024'!B2:B9)+Data!$C$3",
            "=Sales[[#This Row],[Q1 2024]]*Sales[FY2024]",
            "=SUM(Jan:Mar!B2)+[1]Prices!A1+'[Book 2.xlsx]Sheet 1'!C4",
            "=SUM(A:A)-SUM(2:3)+1E5*B2",
            '=IF(B2>0,"see C3",D4)&_xlfn.LET(_xlpm.x,E5,_xlpm.x*2)',
            "=SUM(_xlfn.ANCHORARRAY(F6))+G7#",
            "=Model!#REF!+H8",
            '=INDIRECT("I9")+OFFSET(J10,1,1)',
            "=_A1+XA1*$B$2",
            "=#N/AA1+B2",
        ]
        for formula in tricky:
            for row in (1, 50, 4000):  # the same shape with other cells, read twice to use what was kept
                text = formula.replace("B2", f"B{row}").replace("C3", f"C{row}")
                with self.subTest(text=text):
                    self.assertEqual(formula_refs(text), formula_refs(text))
                    self.assertEqual(formula_refs(text), _tokenised(text))

    def test_formulas_show_as_excel_shows_them(self):
        cases = {
            "=_xlfn.XLOOKUP(A1,B:B,C:C)": "=XLOOKUP(A1,B:B,C:C)",
            "=_xlfn._xlws.SORT(A1:A5)": "=SORT(A1:A5)",
            "=_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)": "=LET(x,A1,x*2)",
            "=SUM(_xlfn.ANCHORARRAY(B2))": "=SUM(B2#)",
            "=SUM(_xlfn.ANCHORARRAY('Q1 data'!$B$2))": "=SUM('Q1 data'!$B$2#)",
            "=_xlfn.SINGLE(A1:A10)*2": "=@A1:A10*2",
            '="_xlfn.X"&A1': '="_xlfn.X"&A1',
            "=B2*(1+Growth)": "=B2*(1+Growth)",
        }
        for stored, shown in cases.items():
            with self.subTest(stored=stored):
                self.assertEqual(excel_formula(stored), shown)


def _tokenised(formula):
    """What formula_refs finds, worked out by tokenising the whole formula."""
    from openpyxl.formula import Tokenizer

    from sheetscope.refs import _SPILL, _STRING

    notes, text = [], _STRING.sub('""', formula)
    if _SPILL.search(text):
        text = _SPILL.sub(r"\1", text)
        notes.append("reads a spilled range (A1#), and only its first cell is known")
    tokens = Tokenizer(text).items
    operands = [t for t in tokens if t.type == "OPERAND"]
    refs = [
        t.value
        for t in operands
        if t.subtype == "RANGE" and not t.value.startswith("_xlpm.") and "#REF!" not in t.value.upper()
    ]
    functions = {t.value.upper() for t in tokens if t.type == "FUNC"}
    if "_XLFN.ANCHORARRAY(" in functions and not notes:
        notes.append("reads a spilled range (A1#), and only its first cell is known")
    if any("#REF!" in t.value.upper() for t in operands):
        notes.append("reads a deleted range (#REF!)")
    if functions & {"INDIRECT(", "OFFSET("}:
        notes.append("builds references with INDIRECT or OFFSET, which are only known when Excel calculates")
    return refs, " and ".join(notes) or None


if __name__ == "__main__":
    unittest.main()
