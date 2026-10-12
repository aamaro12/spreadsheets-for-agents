"""The shown value follows Excel's number formats, so a quote matches what a person sees on screen."""

from datetime import date, datetime, time, timedelta
import unittest

from sheetscope.formats import display

UK_ACCOUNTING = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'


class NumberFormatTests(unittest.TestCase):
    def test_numbers(self):
        cases = [
            (1234.5, "General", "1234.5"),
            (1234, "General", "1234"),
            (0.1 + 0.2, "General", "0.3"),
            (1e20, "General", "1E+20"),
            (1234.567, "0", "1235"),
            (2.675, "0.00", "2.68"),
            (1234567.891, "#,##0", "1,234,568"),
            (1234567.891, "#,##0.00", "1,234,567.89"),
            (-1234.5, "#,##0", "-1,235"),
            (-1234.4, "#,##0;(#,##0)", "(1,234)"),
            (0, '#,##0;(#,##0);"-"', "-"),
            (0.125, "0.0%", "12.5%"),
            (0.1, "0%", "10%"),
            (1234567, '#,##0.0,,"m"', "1.2m"),
            (1234567, "#,##0,", "1,235"),
            (1500, '"£"#,##0', "£1,500"),
            (1500, "[$€-407]#,##0.00", "€1,500.00"),
            (-1500, "[Red]#,##0;[Red](#,##0)", "(1,500)"),
            (12.345, "0.0\\x", "12.3x"),
            (12.345, '0.0"x"', "12.3x"),
            (0.5, "#.00", ".50"),
            (7, "00", "07"),
            (123456, "0.00E+00", "1.23E+05"),
            (0.000123, "0.0E+0", "1.2E-4"),
            (9.999, "0.0E+00", "1.0E+01"),
            (1234.5, "#,##0.00_);(#,##0.00)", "1,234.50"),
            (-1234.5, "#,##0.00_);(#,##0.00)", "(1,234.50)"),
            (1234.5, UK_ACCOUNTING, "1,234.50"),
            (-1234.5, UK_ACCOUNTING, "-1,234.50"),
            (0, UK_ACCOUNTING, "-"),
            (1500, '[>=1000]#,##0,"k";0', "2k"),
            (999, '[>=1000]#,##0,"k";0', "999"),
            (12.5, "0.##", "12.5"),
            (42, "General\\ \\u\\n\\i\\t\\s", "42 units"),
            (1234.5, "@", "1234.5"),
            (1234.5, '"Total: "@', "Total: 1234.5"),
            (101.5, "0 0/32", "101.5"),
            (1.5, "# ?/?", "1.5"),
        ]
        for value, fmt, shown in cases:
            with self.subTest(value=value, fmt=fmt):
                self.assertEqual(display(value, fmt), shown)

    def test_a_number_saved_with_a_date_format_shows_as_a_date(self):
        self.assertEqual(display(45747, "yyyy-mm-dd"), "2025-03-31")


class TextAndDateTests(unittest.TestCase):
    def test_text_booleans_and_empty_cells(self):
        self.assertEqual(display("abc", "@"), "abc")
        self.assertEqual(display("abc", '0;-0;0;"Name: "@'), "Name: abc")
        self.assertEqual(display("abc", "0.0%"), "abc")
        self.assertEqual(display(True, "General"), "TRUE")
        self.assertEqual(display(None, "0.0"), "")

    def test_durations_count_elapsed_time(self):
        self.assertEqual(display(timedelta(hours=26, minutes=30), "[h]:mm"), "26:30")
        self.assertEqual(display(timedelta(minutes=90, seconds=5), "[mm]:ss"), "90:05")
        self.assertEqual(display(timedelta(hours=26, minutes=30), "h:mm"), "2:30")

    def test_dates_and_times(self):
        day = datetime(2025, 3, 31)
        moment = datetime(2025, 3, 31, 14, 5, 9)
        cases = [
            (day, "General", "2025-03-31"),
            (day, "yyyy-mm-dd", "2025-03-31"),
            (day, "mmm-yy", "Mar-25"),
            (day, "dd/mm/yyyy", "31/03/2025"),
            (day, "[$-409]mmmm d, yyyy;@", "March 31, 2025"),
            (day, "dddd d mmm", "Monday 31 Mar"),
            (moment, "dd/mm/yyyy hh:mm", "31/03/2025 14:05"),
            (moment, "h:mm AM/PM", "2:05 PM"),
            (moment, "hh:mm:ss", "14:05:09"),
            (moment, "General", "2025-03-31 14:05:09"),
            (date(2025, 1, 2), "d mmmm yyyy", "2 January 2025"),
            (time(9, 30), "h:mm", "9:30"),
        ]
        for value, fmt, shown in cases:
            with self.subTest(value=value, fmt=fmt):
                self.assertEqual(display(value, fmt), shown)


if __name__ == "__main__":
    unittest.main()
