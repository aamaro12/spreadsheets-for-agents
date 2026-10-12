"""Command line: sheetscope outline | find | read | cell | trace FILE ..., or sheetscope mcp."""

import argparse
import json
import sys

from . import __version__
from .book import FIND_LIMIT, READ_CELLS, TRACE_DEPTH, SheetError, load
from .render import markdown


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="sheetscope",
        description="Read an Excel workbook in steps: outline, find, read, cell and trace. Every value comes "
        "with the cell it came from.",
    )
    parser.add_argument("--version", action="version", version=f"sheetscope {__version__}")
    commands = parser.add_subparsers(dest="command", required=True, metavar="command")

    def step(name, summary):
        command = commands.add_parser(name, help=summary, description=summary)
        command.add_argument("file", help="an .xlsx or .xlsm workbook")
        command.add_argument("--json", action="store_true", help="print JSON instead of Markdown")
        return command

    step("outline", "Sheets, blocks, header rows, row labels, defined names and tables")
    find = step("find", "Cells whose text or shown value contains the query, or whose number rounds to it")
    find.add_argument("query")
    find.add_argument("--sheet", help="search one sheet only")
    find.add_argument("--formulas", action="store_true", help="search formula text instead of values")
    find.add_argument("--limit", type=int, default=FIND_LIMIT, help=f"most matches to show (default {FIND_LIMIT})")
    read = step("read", "A range as a grid, with the labels and headers around it and its formulas")
    read.add_argument("ref", help="a range such as Model!B3:F12, a sheet, a defined name or a table")
    read.add_argument("--max-cells", type=int, default=READ_CELLS, help=f"most cells to show (default {READ_CELLS})")
    cell = step("cell", "One cell in full: value, format, formula, the cells it reads and the formulas that use it")
    cell.add_argument("ref", help="a cell such as Model!D4, a defined name or a citation")
    trace = step("trace", "The tree of cells behind a formula, down to its inputs")
    trace.add_argument("ref", help="a cell such as Summary!B4")
    trace.add_argument("--dependents", action="store_true", help="follow the formulas that use the cell instead")
    trace.add_argument("--depth", type=int, default=TRACE_DEPTH, help=f"levels to follow (default {TRACE_DEPTH})")
    serve = commands.add_parser("mcp", help="Serve the same steps as MCP tools over stdio")
    serve.add_argument("--root", default=".", help="the folder the server may read (default: this folder)")
    args = parser.parse_args(argv)

    try:
        if args.command == "mcp":
            return _serve(args.root)
        book = load(args.file)
        if args.command == "outline":
            result = book.outline()
        elif args.command == "find":
            result = book.find(args.query, sheet=args.sheet, formulas=args.formulas, limit=args.limit)
        elif args.command == "read":
            result = book.read(args.ref, max_cells=args.max_cells)
        elif args.command == "cell":
            result = book.cell(args.ref)
        else:
            result = book.trace(args.ref, dependents=args.dependents, depth=args.depth)
    except SheetError as error:
        print(f"sheetscope: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2) if args.json else markdown(result))
    return 0


def _serve(root) -> int:
    try:
        from .server import serve
    except ModuleNotFoundError as error:
        if (error.name or "").split(".")[0] != "mcp":
            raise
        print("sheetscope: the MCP server needs the mcp extra: pip install 'sheetscope[mcp]'", file=sys.stderr)
        return 2
    serve(root)
    return 0
