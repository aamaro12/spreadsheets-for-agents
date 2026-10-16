"""Row labels and column headers are found the way a person scans a sheet."""

from datetime import datetime
import random
from types import SimpleNamespace
import unittest

from sheetscope.layout import Layout
from sheetscope.refs import Area


def grid(rows, merged=()):
    """A layout over rows of values, where None is an empty cell."""
    cells = {}
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            if value is not None:
                kind = "text" if isinstance(value, str) else "date" if isinstance(value, datetime) else "number"
                cells[(r, c)] = SimpleNamespace(value=value, kind=kind)
    return Layout(cells, merged)


def text(layout, positions):
    """The cells' values joined as the reader joins them, or None."""
    return " / ".join(str(layout.cells[p].value) for p in positions) or None


class ModelTests(unittest.TestCase):
    def setUp(self):
        self.layout = grid(
            [
                ["Northwind: operating model"],
                [],
                ["£k", "FY2024", "FY2025", "FY2026"],
                ["Revenue", 1200, 1320, 1452],
                ["Growth", None, 0.1, 0.1],
                [],
                ["Costs"],
                ["Operating costs", 310, 319, 329],
                [],
                [],
                ["Headcount", 12, 13, 15],
            ]
        )

    def test_a_figure_takes_its_header_row_and_row_label(self):
        layout = self.layout
        self.assertEqual(text(layout, layout.header(4, 3)), "FY2025")
        self.assertEqual(text(layout, layout.label(4, 3)), "Revenue")
        self.assertEqual(text(layout, layout.header(8, 4)), "FY2026")

    def test_a_section_below_a_gap_keeps_the_header_above(self):
        self.assertEqual(text(self.layout, self.layout.header(11, 2)), "FY2024")

    def test_headers_and_titles_have_no_header_of_their_own(self):
        self.assertIsNone(text(self.layout, self.layout.header(3, 2)))
        self.assertIsNone(text(self.layout, self.layout.header(1, 1)))

    def test_blocks_are_cells_within_two_rows_and_one_column_of_each_other(self):
        rng = random.Random(3)
        for _ in range(200):
            spots = {(rng.randint(1, 12), rng.randint(1, 8)) for _ in range(rng.randint(1, 40))}
            layout = grid([[1 if (r, c) in spots else None for c in range(1, 9)] for r in range(1, 13)])
            blocks, seen = [], set()
            for start in sorted(spots):  # the blocks a flood fill finds, in order of their first cell
                if start in seen:
                    continue
                block, stack = set(), [start]
                while stack:
                    r, c = stack.pop()
                    if (r, c) in spots and (r, c) not in block:
                        block.add((r, c))
                        stack += [(r + dr, c + dc) for dr in range(-2, 3) for dc in (-1, 0, 1)]
                seen |= block
                blocks.append(block)
            self.assertEqual([set(layout.members(b)) for b in range(len(layout.areas))], blocks)
            self.assertTrue(all(layout.block[p] == i for i, block in enumerate(blocks) for p in block))

    def test_blocks_bridge_one_blank_row_but_not_two(self):
        self.assertEqual([area.a1() for area in self.layout.areas], ["A1:D8", "A11:D11"])
        self.assertEqual(self.layout.header_rows(0), [3])


class HeaderTests(unittest.TestCase):
    def test_years_and_dates_make_a_header_row(self):
        years = grid([["£k", 2024, 2025], ["Revenue", 1200, 1320], ["Costs", 300, 310]])
        self.assertEqual(text(years, years.header(3, 3)), "2025")
        falling = grid([["£k", 2026, 2025, 2024], ["Revenue", 10, 9, 8]])
        self.assertEqual(text(falling, falling.header(2, 2)), "2026")
        dates = grid([[None, datetime(2024, 12, 31)], ["Revenue", 1200]])
        self.assertEqual(text(dates, dates.header(2, 2)), "2024-12-31 00:00:00")

    def test_figures_that_look_like_years_are_not_a_header(self):
        layout = grid(
            [["£k", "North", "South", "West"], ["Revenue", 1950, 2010, 2080], ["Costs", 1250.25, 1300.5, 1410.75]]
        )
        self.assertEqual(text(layout, layout.header(3, 3)), "South")
        self.assertEqual(text(layout, layout.header(2, 2)), "North")

    def test_text_among_figures_is_not_a_header(self):
        layout = grid([["Company", "P/E", "EV/EBITDA"], ["Alder", "NM", 8.1], ["Birch", 14.2, 9.3]])
        self.assertEqual(text(layout, layout.header(3, 2)), "P/E")

    def test_stacked_header_rows_are_joined_across_merged_cells(self):
        layout = grid(
            [
                [None, "FY2024", None, "FY2025", None],
                ["£k", "Actual", "Budget", "Actual", "Budget"],
                ["Revenue", 100, 110, 120, 125],
            ],
            merged=[Area(1, 2, 1, 3), Area(1, 4, 1, 5)],
        )
        self.assertEqual(text(layout, layout.header(3, 3)), "FY2024 / Budget")
        self.assertEqual(text(layout, layout.header(3, 4)), "FY2025 / Actual")

    def test_a_label_value_unit_list_keeps_its_header_row(self):
        layout = grid([["Assumption", "Value", "Unit"], ["Growth", 0.1, "a year"], ["Tax rate", 0.25, "of profit"]])
        self.assertEqual(text(layout, layout.header(3, 2)), "Value")
        self.assertEqual(text(layout, layout.header(3, 3)), "Unit")

    def test_placeholders_such_as_nm_are_values_not_headings(self):
        layout = grid([["£m", "FY2024", "FY2025", "FY2026"], ["Growth", "NM", "n/a", "–"], ["Revenue", 100, 110, 121]])
        self.assertEqual(text(layout, layout.header(3, 3)), "FY2025")
        no_header = grid([["Revenue", 100, 110], ["Costs", 60, 70], ["Margin", "NM", "NM"], ["Profit", 40, 40]])
        self.assertIsNone(text(no_header, no_header.header(4, 2)))

    def test_a_units_row_under_the_years_joins_their_header(self):
        units = grid([[None, "FY2024", "FY2025"], [None, "£m", "£m"], ["Revenue", 100, 110]])
        self.assertEqual(text(units, units.header(3, 2)), "FY2024 / £m")

    def test_values_may_start_a_few_rows_below_their_header(self):
        scenarios = grid(
            [
                ["USD'000", None, "Base", "Bear", "Bull"],
                ["Cash", 120, "✓", "✓", "✓"],
                ["Renewals", 80, "✓", None, "✓"],
                ["New contracts", 40, None, None, "✓"],
                [],
                ["Monthly spend", None, 30, 35, 28],
            ]
        )
        self.assertEqual((scenarios.header_rows(0), text(scenarios, scenarios.header(6, 4))), ([1], "Bear"))
        errors = grid([[None, "Price > $75", "Price < $45"], *[[None, "#N/A", "#N/A"]] * 12, [None, 0.2, 0.1]])
        for cell in errors.cells.values():
            cell.kind = "error" if cell.value == "#N/A" else cell.kind
        self.assertEqual((errors.header_rows(0), text(errors, errors.header(14, 3))), ([1], "Price < $45"))

    def test_text_below_figures_does_not_start_a_header(self):
        layout = grid([["£m", "FY2024", "FY2025"], ["Revenue", 100, 110], ["Rating", "AA", "A"], ["Costs", 60, 70]])
        self.assertEqual(text(layout, layout.header(4, 3)), "FY2025")
        untitled = grid([["Revenue", 100, 110], ["Note", "see below", None], ["Costs", 60, 70]])
        self.assertIsNone(text(untitled, untitled.header(3, 2)))

    def test_rows_of_labels_and_text_values_are_not_headers(self):
        layout = grid(
            [
                ["Entity information", "12 Months Ended", None],
                [None, "Dec. 31, 2023", "Jun. 30, 2023"],
                ["Document type", "10-K", None],
                ["Shell company", "true", None],
                ["Public float", None, 8018000],
            ]
        )
        self.assertEqual(layout.header_rows(0), [])
        self.assertEqual(text(layout, layout.header(5, 3)), "Jun. 30, 2023")
        self.assertEqual(text(layout, layout.label(5, 3)), "Public float")

    def test_a_label_and_its_text_value_inside_a_table_are_not_a_header(self):
        terms = grid(
            [
                [None, "Dec. 31, 2023", "Dec. 31, 2022"],
                ["Warrants outstanding", 100, 90],
                ["Exercise period", "30 days", None],
                ["Expiration", "5 years", None],
                ["Redemption price", 18, None],
            ]
        )
        self.assertEqual(text(terms, terms.header(5, 2)), "Dec. 31, 2023")
        cash = grid(
            [["Month", "Revenue", "Costs"], ["Jan", 10, 6], ["Feb", 11, 7], [], ["Month", "Ending cash"], ["Jan", 40]]
        )
        self.assertEqual(text(cash, cash.header(6, 2)), "Ending cash")

    def test_a_header_band_is_three_rows_at_most(self):
        layout = grid(
            [[None, "a", "b"], [None, "c", "d"], [None, "e", "f"], ["£m", "FY2024", "FY2025"], ["Revenue", 100, 110]]
        )
        self.assertEqual(text(layout, layout.header(5, 3)), "d / f / FY2025")

    def test_settings_are_not_headers(self):
        above = grid(
            [["Scenario", "Base"], ["Currency", "USD ($000)"], ["Item", "Q1", "Q2"], ["Opening cash", 100, 90]]
        )
        self.assertEqual(text(above, above.header(4, 3)), "Q2")
        settings = grid(
            [["Company", "Acme"], ["Currency", "USD"], ["Units", "m"], ["Tax rate", 0.25], ["Growth", 0.03]]
        )
        self.assertEqual((settings.header_rows(0), text(settings, settings.header(4, 2))), ([], None))
        table = grid([["Item", "Amount"], ["Rent", 100], ["Rates", 40]])
        self.assertEqual(text(table, table.header(3, 2)), "Amount")

    def test_a_label_and_value_header_stops_at_a_blank_row(self):
        inputs = grid(
            [
                ["Scenario", "Revenue growth"],
                ["Base", 0.06],
                ["Upside", 0.11],
                [],
                ["Gross margin", 0.38],
                ["Tax rate", 0.25],
            ]
        )
        self.assertEqual(text(inputs, inputs.header(3, 2)), "Revenue growth")
        self.assertIsNone(text(inputs, inputs.header(5, 2)))
        self.assertIsNone(text(inputs, inputs.header(6, 2)))

    def test_a_header_row_apart_from_its_figures_still_heads_them(self):
        months = [datetime(2024, m, 28) for m in (1, 2, 3)]
        apart = grid([[None, *months], [], [], ["Revenue", 10, 11, 12], ["Costs", 5, 6, 7]])
        self.assertEqual(text(apart, apart.header(5, 3)), "2024-02-28 00:00:00")
        counted = grid([[None, 0, 1, 2], [None, *months], [], [], ["Revenue", 10, 11, 12]])
        self.assertEqual(text(counted, counted.header(5, 4)), "2024-03-28 00:00:00")
        customers = grid([["Customer", "Region"], ["Alder", "North"], [], [], ["Revenue", 10, 11]])
        self.assertIsNone(text(customers, customers.header(5, 2)))
        timeline = grid(
            [
                ["Northwind: model"],
                [],
                ["Scenario", "Base"],
                ["£m", "FY2024", "FY2025"],
                [],
                [],
                [],
                ["Revenue", 10, 11],
            ]
        )
        self.assertEqual(text(timeline, timeline.header(8, 3)), "FY2025")

    def test_figures_scattered_below_a_header_row_take_it(self):
        filing = grid(
            [
                ["Related party", None, None, "12 Months Ended"],
                [None, "Dec. 31, 2023", "Jun. 09, 2023", "Dec. 31, 2022", "Jan. 26, 2023", "Jul. 27, 2023"],
                ["Commercial paper"],
                ["Related party"],
                ["Interest rate", None, None, None, None, 0.05],
                ["Interest expense", 120],
                ["Long-term debt"],
                ["Debt facility", 1500, None, 900, None, 1100],
            ],
            merged=[Area(1, 4, 1, 6)],
        )
        self.assertEqual(filing.header_rows(0), [2])
        self.assertEqual(text(filing, filing.header(8, 4)), "12 Months Ended / Dec. 31, 2022")
        titled = grid(
            [["Entity information", "12 Months Ended"], [None, "Dec. 31, 2023", "Dec. 31, 2022"], ["Float", 80, 70]],
            merged=[Area(1, 1, 2, 1)],
        )
        self.assertEqual(text(titled, titled.header(3, 2)), "12 Months Ended / Dec. 31, 2023")
        shares = grid(
            [
                ["Shareholder", "Ordinary", None, None, "Issued"],
                [None, "Options", "Shares", None, "Shares"],
                [],
                ["Founders"],
                ["A. Founder", None, 4000, None, 4000],
            ],
            merged=[Area(1, 2, 1, 3)],
        )
        self.assertEqual(text(shares, shares.header(5, 3)), "Ordinary / Shares")

    def test_a_mostly_text_row_heads_a_table_without_a_header_row(self):
        layout = grid([[None, "Code", "Description", "Constant", "Lags", 2], ["x", "S1", "Farm", 0.5, None, 0.1]])
        self.assertEqual(text(layout, layout.header(2, 4)), "Constant")
        titled = grid(
            [["Income statement"], ["Revenue", 100, 110, "note"], ["Costs", 60, 70, "note"]], [Area(1, 1, 1, 4)]
        )
        self.assertIsNone(text(titled, titled.header(3, 3)))

    def test_a_count_below_a_timeline_is_data(self):
        dcf = grid(
            [
                ["£m", 2024, 2025, 2026],
                ["Free cash flow", 100, 110, 121],
                ["Discount period", 1, 2, 3],
                ["Discount factor", 0.95, 0.9, 0.86],
            ]
        )
        self.assertEqual(text(dcf, dcf.header(4, 3)), "2025")
        months = [datetime(2025, m, 28) for m in (1, 2, 3)]
        monthly = grid([["Month", *months], ["Sales", 1, 2, 3], ["Costs", 5, 6, 7]])
        self.assertEqual(
            (text(monthly, monthly.header(3, 2)), text(monthly, monthly.label(2, 2))), ("2025-01-28 00:00:00", "Sales")
        )

    def test_years_may_step_by_more_than_one(self):
        layout = grid([["Mt CO2", 2025, 2030, 2040, 2050], ["Emissions", 40, 31, 18, 5]])
        self.assertEqual(text(layout, layout.header(2, 4)), "2040")

    def test_a_count_beside_text_headings_heads_its_columns(self):
        layout = grid(
            [
                [None, None, None, None, None, "AR coefficients"],
                ["Code", "Description", None, "Constant", "Lag", 1, 2, 3, 4],
                [None, "Farm", "SFZ", -2.7, None, 0.44, 0.21, -0.13, -0.26],
                [None, "Mining", "SNOZ", 0.16, None, 0.47, 0.07, -0.12, 0.09],
            ]
        )
        self.assertEqual(layout.header_rows(0), [2])
        self.assertEqual((text(layout, layout.header(4, 4)), text(layout, layout.header(4, 7))), ("Constant", "2"))
        upper = grid(
            [
                [None, None, "Constant", None, "Lags"],
                ["Code", "Name", None, "Lag", 1, 2, 3],
                ["S1", "Farm", 0.5, None, 0.4, 0.2, 0.1],
            ]
        )
        self.assertEqual(text(upper, upper.header(3, 3)), "Constant")

    def test_a_table_of_text_takes_its_own_first_row(self):
        layout = grid([["Customer", "Region"], ["Alder Foods", "North"], ["Birch Retail", "South"]])
        self.assertEqual(text(layout, layout.header(3, 2)), "Region")
        below = grid(
            [
                ["£k", "FY2024", "FY2025"],
                ["Revenue", 100, 110],
                [],
                [],
                [],
                ["Customer", "Region"],
                ["Alder", "North"],
                ["Birch", "South"],
            ]
        )
        self.assertEqual(text(below, below.header(8, 2)), "Region")


class LabelTests(unittest.TestCase):
    def test_labels_join_the_text_columns_to_the_left(self):
        layout = grid([["Region", "Metric", "FY2024"], ["UK", "Revenue", 100], ["UK", "Costs", 60]])
        self.assertEqual(text(layout, layout.label(3, 3)), "UK / Costs")
        units = grid([[None, None, "FY2024"], ["Revenue", "£k", 1200]])
        self.assertEqual(text(units, units.label(2, 3)), "Revenue / £k")
        assumptions = grid(
            [["Assumption", "Value", "Unit"], ["Growth", 0.1, "a year"], ["Tax rate", 0.25, "of profit"]]
        )
        self.assertEqual(text(assumptions, assumptions.label(3, 3)), "Tax rate")

    def test_text_among_figures_is_not_a_label(self):
        layout = grid([["Company", "P/E", "EV/EBITDA"], ["Alder", "NM", 8.1], ["Birch", 14.2, 9.3]])
        self.assertEqual(text(layout, layout.label(2, 3)), "Alder")

    def test_labels_come_from_the_same_table(self):
        layout = grid(
            [
                ["Metric", "FY2025", None, "Region", "Sales"],
                ["Revenue", 1320, None, "North", 700],
                ["Costs", 319, None, "South", 620],
            ]
        )
        self.assertEqual(text(layout, layout.label(2, 5)), "North")
        self.assertEqual(text(layout, layout.label(2, 2)), "Revenue")
        self.assertEqual(text(layout, layout.header(3, 5)), "Sales")

    def test_an_empty_label_is_not_taken_from_another_table(self):
        layout = grid(
            [
                ["Scheme", "Value", None, "Item", "Amount"],
                ["Pensions", 10, None, "Rent", 5],
                ["Bonds", 20, None, None, 7],
            ]
        )
        self.assertEqual(text(layout, layout.label(2, 5)), "Rent")
        self.assertIsNone(text(layout, layout.label(3, 5)))

    def test_figures_across_a_spacer_column_take_the_labels_beside_them(self):
        layout = grid([[None, None, "FY2024", "FY2025"], ["Revenue", None, 1200, 1320]])
        self.assertEqual(text(layout, layout.label(2, 4)), "Revenue")

    def test_notes_to_the_right_of_figures_are_not_their_labels(self):
        layout = grid(
            [
                ["Revenue", None, None, 100, 110, "on plan", "ok"],
                ["Costs", None, None, 60, 70, "late", "ok"],
                ["Profit", None, None, 40, 40, "flat", "ok"],
            ]
        )
        self.assertEqual(text(layout, layout.label(2, 5)), "Costs")
        self.assertEqual(text(layout, layout.label(3, 6)), "Profit")

    def test_a_table_led_by_dates_or_years_takes_them_as_row_labels(self):
        prices = grid(
            [["Date", "Open", "Close"], [datetime(2024, 1, 2), 10.0, 10.5], [datetime(2024, 1, 3), 10.5, 10.2]]
        )
        self.assertEqual(text(prices, prices.label(3, 3)), "2024-01-03 00:00:00")
        years = grid([["Year", "Revenue"], [2023, 100], [2024, 110]])
        self.assertEqual(text(years, years.label(3, 2)), "2024")
        codes = grid([["Code", "Amount"], [4010, 100], [4020, 110]])
        self.assertIsNone(text(codes, codes.label(3, 2)))
        runway = grid(
            [
                ["Capital raised", None, None, 3250],
                ["Cash available", None, None, 2],
                ["Runway", None, datetime(2026, 8, 23), 4],
            ]
        )
        self.assertEqual(text(runway, runway.label(1, 4)), "Capital raised")

    def test_a_label_that_repeats_takes_the_heading_of_its_section(self):
        filing = grid(
            [
                ["Revenue by region (Details)", "FY2024", "FY2023"],
                ["Revenues"],
                ["Revenue", 60922, 26974],
                ["United States"],
                ["Revenues"],
                ["Revenue", 26966, 8292],
                ["Taiwan"],
                ["Revenues"],
                ["Revenue", 13405, 6986],
                ["Operating income", 32972, 4224],
            ]
        )
        expected = {3: "Revenue", 6: "United States / Revenue", 9: "Taiwan / Revenue", 10: "Operating income"}
        for row, label in expected.items():
            with self.subTest(row=row):
                self.assertEqual(text(filing, filing.label(row, 2)), label)
        accounts = grid(
            [
                ["£k", "FY2024"],
                ["Revenue"],
                ["Product", 100],
                ["Services", 50],
                ["Total", 150],
                ["Costs"],
                ["Staff", 60],
                ["Total", 60],
            ]
        )
        self.assertEqual(
            [text(accounts, accounts.label(r, 2)) for r in (3, 5, 8)], ["Product", "Revenue / Total", "Costs / Total"]
        )

    def test_a_table_led_by_dates_and_text_takes_both_as_row_labels(self):
        sales = grid(
            [
                ["Month", "Region", "Revenue"],
                [datetime(2024, 1, 1), "North", 310],
                [datetime(2024, 1, 1), "South", 265],
                [datetime(2024, 2, 1), "North", 320],
            ]
        )
        self.assertEqual(text(sales, sales.label(4, 3)), "2024-02-01 00:00:00 / North")
        self.assertEqual(text(sales, sales.header(4, 3)), "Revenue")

    def test_labels_above_a_table_further_down_the_block_are_kept(self):
        rates = grid(
            [
                ["GBP-USD FX", 1.27, "Update monthly"],
                ["EUR-USD FX", 1.08, "Update monthly"],
                ["Source"],
                ["Owner"],
                ["Reviewer"],
                ["Rate", "Amount", "Date"],
                [0.05, 100, datetime(2026, 1, 31)],
                [0.06, 110, datetime(2026, 2, 28)],
            ]
        )
        self.assertEqual(text(rates, rates.label(2, 2)), "EUR-USD FX")


if __name__ == "__main__":
    unittest.main()
