"""Where a sheet's labels are: blocks of cells, header rows, row labels and column headers.

These are reading aids, not facts stored in the file. A row label is the text in its block's
leading text columns, to the left of the cell, after the date or year that leads its row; a label
that repeats in its block, such as Revenue under each region of a filing, starts with the heading
of its section. A column header comes from the nearest header row above: text, dates or consecutive
years with values below them, joined with the header rows stacked directly above it. The header of
a list of labels and values stops at a blank row. Placeholders such as NM and n/a are values, not
headings. A list of text alone takes its own first row.
"""

from bisect import bisect_left
from collections import Counter, defaultdict

from .refs import Area

MAX_MERGE, MAX_ANCHORED = 10_000, 500_000  # cells of one merged range, and of all, that heading lookups cover
MAX_BAND = 3  # header rows joined into one column header
REACH = 10  # rows below a header row within which its values start
SAMPLE = 2_000  # rows of a block counted to find its label columns
PLACEHOLDERS = {"nm", "n/m", "n.m.", "nmf", "na", "n/a", "n.a.", "-", "--", "–", "—"}
VALUES = ("number", "date", "error", "boolean")


class Layout:
    def __init__(self, cells, merged=()):
        """cells maps (row, col) to objects with a kind ("text", "number", "date", ...) and a value."""
        self.cells = cells
        self.rows = defaultdict(list)
        for row, col in sorted(cells):
            self.rows[row].append(col)
        self.anchor = {}
        if cells and merged:
            used = Area(min(self.rows), min(c for _, c in cells), max(self.rows), max(c for _, c in cells))
            for area in merged:
                inside = area.clip(used)
                if inside and inside.size <= MAX_MERGE and len(self.anchor) + inside.size <= MAX_ANCHORED:
                    corner = (area.min_row, area.min_col)
                    for r in range(inside.min_row, inside.max_row + 1):
                        for c in range(inside.min_col, inside.max_col + 1):
                            self.anchor[(r, c)] = corner
        self.block = {}
        self.areas: list[Area] = []
        self._runs: list[list[tuple[int, int, int]]] = []  # each block's runs of adjacent cells in a row
        self._rows: list[list[int]] = []
        self._row_sets: list[set[int]] = []
        self._heads, self._labels, self._figures, self._above, self._like = {}, {}, {}, {}, {}
        self._sections: dict[int, dict[int, int]] = {}  # each block's rows that take a section heading
        self._keyed: set[int] = set()  # blocks whose rows are labelled by a leading date or year
        self._dated_from: dict[int, int | None] = {}  # each block's first row of dates or years
        self._group()

    def _group(self):
        """Group cells into blocks: neighbours, allowing 1 blank row but no blank column between them.
        Each row's runs of adjacent cells join the runs they touch in the 2 rows above."""
        runs, spans = [], {}  # runs of adjacent cells as (row, first col, last col), in reading order
        for row in sorted(self.rows):
            cols, start = self.rows[row], len(runs)
            first = cols[0]
            for previous, col in zip(cols, cols[1:]):
                if col != previous + 1:
                    runs.append((row, first, previous))
                    first = col
            runs.append((row, first, cols[-1]))
            spans[row] = range(start, len(runs))
        parent = list(range(len(runs)))

        def root(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        for row, here in spans.items():
            for above in (spans.get(row - 1), spans.get(row - 2)):
                i = j = 0
                while above and i < len(here) and j < len(above):
                    _, low, high = runs[here[i]]
                    _, over_low, over_high = runs[above[j]]
                    if over_high < low - 1:
                        j += 1
                    elif high < over_low - 1:
                        i += 1
                    else:
                        a, b = root(here[i]), root(above[j])
                        parent[max(a, b)] = min(a, b)
                        i, j = (i + 1, j) if high < over_high else (i, j + 1)
        index = {}
        for k, (row, low, high) in enumerate(runs):
            block = index.setdefault(root(k), len(index))
            if block == len(self.areas):
                self.areas.append(Area(row, low, row, high))
                self._runs.append([])
                self._rows.append([])
            area = self.areas[block]
            self.areas[block] = Area(area.min_row, min(area.min_col, low), row, max(area.max_col, high))
            if self._rows[block][-1:] != [row]:
                self._rows[block].append(row)
            self._runs[block].append((row, low, high))
            for col in range(low, high + 1):
                self.block[(row, col)] = block
        self._row_sets = [set(rows) for rows in self._rows]

    def rows_of(self, block: int) -> list[int]:
        return self._rows[block]

    def members(self, block: int):
        """The block's cells, row by row."""
        return ((row, col) for row, low, high in self._runs[block] for col in range(low, high + 1))

    def size(self, block: int) -> int:
        return sum(high - low + 1 for _, low, high in self._runs[block])

    def header_rows(self, block: int) -> list[int]:
        """Rows of headings directly above a row that is not headings, with values starting below them."""
        if block not in self._heads:
            rows, heads = self._rows[block], []
            like = [self._heading_row(block, r) for r in rows]
            for i, r in enumerate(rows):
                if like[i] and i + 1 < len(rows) and not like[i + 1] and self._over_values(block, r):
                    if self._pair(block, r) and self._pair(block, r - 1):
                        continue  # a list of settings such as Currency | USD, not a table
                    narrow = len(self._in_row(block, r)) < 3 and r - 1 in self._row_sets[block]
                    if not heads or not narrow:  # later, 2 cells straight after a row are a label and its value
                        heads.append(r)
            self._heads[block] = heads
        return self._heads[block]

    def band(self, block: int, top: int) -> list[int]:
        """A header row and the header rows stacked directly above it, top first. A setting above a table,
        such as Currency | USD, is a label and its value rather than a header."""
        rows = [top]
        while len(rows) < MAX_BAND and rows[0] - 1 in self._row_sets[block]:
            above = rows[0] - 1
            if not self._heading_row(block, above) or self._pair(block, above):
                break
            rows.insert(0, above)
        return rows

    def headings(self, block: int, top: int, col: int) -> list[tuple[int, int]]:
        """The cells heading a column in the header band that ends at row top, top first."""
        found = []
        for row in self.band(block, top):
            if (position := self._heading_at(row, col)) and position not in found:
                found.append(position)
        return found

    def header(self, row: int, col: int) -> list[tuple[int, int]]:
        """The cells heading this cell's column, top first; empty when there are none."""
        block = self.block.get((row, col))
        if block is None:
            return []
        heads = self.header_rows(block)
        for top in reversed(heads):
            if top < row and (found := self.headings(block, top, col)):
                if len(self._in_row(block, top)) == 2 and self._gap(block, top, row):
                    return []  # the header of a list of labels and values heads it down to a blank row
                return found
        if heads:
            return []
        if self._has_figures(block) and (above := self._header_above(block, col)):
            return above
        for r in self._rows[block]:  # no header row: the first row above that is headings, or mostly text
            if r >= row:
                break
            if self._pair(block, r) and self._has_figures(block):
                continue  # a setting such as Currency | USD above settings with figures
            if self._heading_row(block, r) and self._heading_at(r, col):
                return self.headings(block, r, col)
            if self._text((r, col), "date") and self._wordy(block, r):
                return [(r, col)]
        return []

    def label(self, row: int, col: int, block: int | None = None) -> list[tuple[int, int]]:
        """The cells labelling this cell's row: its block's leading text columns, left of the cell."""
        if block is None:
            block = self.block.get((row, col))
        if block is not None and (columns := self.label_columns(block)):
            kinds = ("date", "number") if block in self._keyed else ()
            found = [(row, c) for c in columns if c < col and self._text((row, c), *kinds)][:3]
            if found and (heading := self.sections(block).get(row)):
                found.insert(0, (heading, columns[0]))
            return found
        limit = self.areas[block].min_col if block is not None else col
        texts = [c for c in self.rows.get(row, ()) if c < limit and self._text((row, c))]
        if not texts:
            return []
        nearest = self.block[(row, texts[-1])]
        return [(row, c) for c in self.label_columns(nearest) or [texts[-1]] if self._text((row, c))][:3]

    def sections(self, block: int) -> dict[int, int]:
        """Rows whose label repeats in the block, each with the heading of its section: the nearest row above
        with a label and no values, whose text is not itself repeated (as a filing repeats "Revenues")."""
        if block not in self._sections:
            columns, found = self.label_columns(block), {}
            if columns and block not in self._keyed:
                first, cells, rest = columns[0], self.cells, columns[1:]
                banded = {r for top in self.header_rows(block) for r in self.band(block, top)}
                keyed = []  # (row, its label, whether it holds values)
                for r in self._rows[block]:
                    cell = cells.get((r, first))
                    if r in banded or cell is None or cell.kind != "text" or (key := _key(cell)) in PLACEHOLDERS:
                        continue
                    label = (key, *(_key(cells.get((r, c))) for c in rest)) if rest else key
                    keyed.append((r, label, self._valued(block, r)))
                repeats = Counter(label for _, label, valued in keyed if valued)
                if any(count > 1 for count in repeats.values()):
                    headings = Counter(label for _, label, valued in keyed if not valued)
                    heading = None
                    for r, label, valued in keyed:
                        if not valued:
                            if headings[label] == 1:
                                heading = r
                        elif heading is not None and repeats[label] > 1:
                            found[r] = heading
            self._sections[block] = found
        return self._sections[block]

    def _valued(self, block, row):
        return any(self.block[(row, c)] == block and self.cells[(row, c)].kind in VALUES for c in self.rows[row])

    def _gap(self, block, top, row):
        """Whether a blank row lies between two rows of a block."""
        rows = self._rows[block]
        return bisect_left(rows, row) - bisect_left(rows, top) != row - top

    def label_columns(self, block: int) -> list[int]:
        """The block's first run of columns holding more text than values, outside its header rows and
        before its values; or, when there is none, a first column of dates or years."""
        if block not in self._labels:
            banded = {r for top in self.header_rows(block) for r in self.band(block, top)}
            rows = [r for r in self._rows[block] if r not in banded]
            counts = defaultdict(lambda: [0, 0])
            for r in rows[:: max(1, len(rows) // SAMPLE)]:
                for c in self.rows[r]:
                    if self.block[(r, c)] == block and self.cells[(r, c)].kind != "empty":
                        counts[c][not self._text((r, c))] += 1
            run, values_before = [], False
            for c in sorted(counts):
                text, values = counts[c]
                if text > values:
                    run.append(c)
                elif values and run:
                    break
                elif values:
                    values_before = True
            else:
                if values_before:
                    run = []  # text after all the values is notes, not labels
            lead = min(counts) if counts else None
            if run and run[0] == lead + 1 and self._key_column(block, lead, banded):
                run.insert(0, lead)  # a date before the text, as in Month | Region | Revenue
                self._keyed.add(block)
            elif not run and counts and self._key_column(block, lead, banded):
                run = [lead]
                self._keyed.add(block)
            self._labels[block] = run
        return self._labels[block]

    def _key_column(self, block, col, banded):
        """A column of dates, or of years one apart, that leads a table's rows: its values label them."""
        rows = [r for r in self._rows[block] if r not in banded]
        values = [self.cells[(r, col)] for r in rows if (r, col) in self.cells]
        if 10 * len(values) < 9 * len(rows):
            return False
        kinds = Counter(cell.kind for cell in values)
        if kinds["date"] > kinds["text"] and not kinds["number"]:
            return True
        years = [int(cell.value) for cell in values if cell.kind == "number" and _year(cell.value)]
        return len(years) >= 2 and len(years) == kinds["number"] > kinds["text"] and _counting(years)

    def _has_figures(self, block):
        if block not in self._figures:
            kinds = (self.cells[p].kind for p in self.members(block))
            self._figures[block] = any(kind in ("number", "date") for kind in kinds)
        return self._figures[block]

    def _header_above(self, block, col):
        """For a block without header rows: the column's header in the nearest header row above the block,
        or in a timeline apart from it at the top of the sheet."""
        if (block, col) not in self._above:
            found = []
            for r in range(self.areas[block].min_row - 1, 0, -1):
                if position := self._heading_at(r, col):
                    other = self.block[position]
                    if r in self.header_rows(other) or self._timeline(other, r):
                        found = self.headings(other, r, col)
                        break
            self._above[(block, col)] = found
        return self._above[(block, col)]

    def _heading_row(self, block, row):
        """A row of 2 or more headings: text, dates, or years that rise or fall, and not a row of text below
        figures, such as ratings."""
        if (block, row) not in self._like:
            self._like[(block, row)] = self._headings_only(block, row)
        return self._like[(block, row)]

    def _headings_only(self, block, row):
        positions, texts, numbers = [], 0, []
        for c in self.rows.get(row, ()):
            if self.block[(row, c)] != block:
                continue
            cell = self.cells[(row, c)]
            if cell.kind == "number":
                if not float(cell.value).is_integer() or numbers and not 1 <= abs(cell.value - numbers[-1]) <= 10:
                    return False  # neither years nor a count
                numbers.append(cell.value)
            elif self._text((row, c)):
                texts += 1
            elif cell.kind != "date":
                return False
            positions.append((row, c))
        if len(positions) < 2 or numbers and not _years(numbers) and not self._count(block, row, numbers, texts):
            return False
        return not self._below_figures(block, row, positions)

    def _count(self, block, row, numbers, texts):
        """A count from 0 or 1 beside text headings, such as lags 1, 2, 3, that heads columns rather than
        counting periods under a timeline, as a DCF's discount periods do."""
        if not (texts and len(numbers) >= 3 and numbers[0] in (0, 1) and _counting(numbers)):
            return False
        if block not in self._dated_from:
            dated = (r for r in self._rows[block] if self._dates_or_years(block, r) >= 2)
            self._dated_from[block] = next(dated, None)
        first = self._dated_from[block]
        return first is None or first > row

    def _dates_or_years(self, block, row):
        cells = (self.cells[p] for p in self._in_row(block, row))
        return sum(cell.kind == "date" or cell.kind == "number" and _year(cell.value) for cell in cells)

    def _wordy(self, block, row):
        """A row holding more text or dates than anything else, not below figures: a header when there is
        no other."""
        positions = self._in_row(block, row)
        headings = [position for position in positions if self._text(position, "date")]
        return 2 * len(headings) > len(positions) and not self._below_figures(block, row, positions)

    def _below_figures(self, block, row, positions):
        """Whether the row's text after its first cell sits mostly directly below a row of figures."""
        if not self._figure_row(block, row - 1):
            return False
        texts = [(r, c) for r, c in positions[1:] if self._text((r, c))]
        under = [(r, c) for r, c in texts if (r - 1, c) in self.cells and self.cells[(r - 1, c)].kind == "number"]
        return 2 * len(under) > len(texts)

    def _figure_row(self, block, row):
        """Whether a row holds figures: numbers other than a count up or down by 1, such as 1, 2, 3 or years."""
        numbers = [self.cells[p].value for p in self._in_row(block, row) if self.cells[p].kind == "number"]
        return bool(numbers) and not (len(numbers) > 1 and (_counting(numbers) or _years(numbers)))

    def _over_values(self, block, row):
        """Whether values start within a few rows below any of this row's cells."""
        for _, c in self._in_row(block, row):
            for r in range(row + 1, row + REACH + 1):
                below = self.cells.get((r, c))
                if below is not None and below.kind in VALUES:
                    return True
                if below is not None and self._heading_row(self.block[(r, c)], r):
                    break
        return False

    def _timeline(self, block, row):
        """A row of headings at the foot of its block, apart from the figures below it: dates or years, or text
        under nothing but a title and settings such as Currency | GBP. The foot of a table of text is not one."""
        if row != self._rows[block][-1] or not self._heading_row(block, row) or self._pair(block, row):
            return False
        rest = [r for r in self._rows[block] if r != row]
        return self._dated(block, row) or all(len(self._in_row(block, r)) < 2 or self._pair(block, r) for r in rest)

    def _dated(self, block, row):
        return any(self.cells[p].kind in ("date", "number") for p in self._in_row(block, row))

    def _pair(self, block, row):
        """A label in the block's first column and one value beside it, neither merged with other cells."""
        positions = self._in_row(block, row)
        if len(positions) != 2 or positions[0][1] != self.areas[block].min_col:
            return False
        return not any(
            self.anchor.get((r + 1, c)) == (r, c) or self.anchor.get((r, c + 1)) == (r, c) for r, c in positions
        )

    def _in_row(self, block, row):
        return [(row, c) for c in self.rows.get(row, ()) if self.block[(row, c)] == block]

    def _heading_at(self, row, col):
        """The cell in a header row that heads a column: the cell itself, or the merged cell over it."""
        position = (row, col) if (row, col) in self.cells else self.anchor.get((row, col))
        return position if position in self.cells else None

    def _text(self, position, *kinds):
        """Whether a cell holds text other than a placeholder such as NM, or a value of the kinds given."""
        cell = self.cells.get(position)
        if cell is None:
            return False
        if cell.kind == "text":
            return str(cell.value).strip().casefold() not in PLACEHOLDERS
        return cell.kind in kinds


def _key(cell):
    return str(cell.value).strip().casefold() if cell is not None else ""


def _year(value):
    return float(value).is_integer() and 1900 <= value <= 2100


def _years(numbers):
    """2 or more years that rise or fall, by up to 10 at a time: 2024, 2025, 2026 or 2030, 2040, 2050."""
    steps = [later - earlier for earlier, later in zip(numbers, numbers[1:])]
    rising, falling = all(1 <= step <= 10 for step in steps), all(-10 <= step <= -1 for step in steps)
    return len(numbers) > 1 and all(_year(number) for number in numbers) and (rising or falling)


def _counting(numbers):
    """Whole numbers that rise or fall by 1, such as 1, 2, 3 or 2026, 2025."""
    steps = {later - earlier for earlier, later in zip(numbers, numbers[1:])}
    return all(float(n).is_integer() for n in numbers) and (steps <= {1} or steps <= {-1})
