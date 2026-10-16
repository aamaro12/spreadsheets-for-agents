"""The MCP server offers the same steps as read-only tools, inside one folder."""

import asyncio
from pathlib import Path
import shutil
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client
from mcp.server.mcpserver.exceptions import ToolError
from openpyxl import Workbook

from sheetscope import load
from sheetscope.server import build_server
from support import northwind


class ServerTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.base = Path(folder.name)
        self.root = self.base / "shared"
        self.root.mkdir()
        northwind.build(self.root / "northwind.xlsx")
        northwind.build(self.base / "private.xlsx")
        self.server = build_server(self.root)

    async def call(self, tool, **arguments):
        result = await self.server.call_tool(tool, arguments)
        return result.content[0].text

    async def test_tools_are_read_only_and_described(self):
        tools = {tool.name: tool for tool in await self.server.list_tools()}
        self.assertEqual(set(tools), {"workbooks", "outline", "find", "read", "cell", "trace"})
        for tool in tools.values():
            self.assertTrue(tool.annotations.read_only_hint, tool.name)
            self.assertTrue(tool.description, tool.name)

    async def test_each_step_returns_markdown(self):
        cases = {
            "workbooks": ({}, "- northwind.xlsx ("),
            "outline": ({"path": "northwind.xlsx"}, "## Model · A1:E12"),
            "find": ({"path": "northwind.xlsx", "query": "1,597"}, "| Model!E4 | 1,597 |"),
            "read": ({"path": "northwind.xlsx", "ref": "Model!C4:E4"}, "- C4:E4: =B4*(1+Growth)"),
            "cell": ({"path": "northwind.xlsx", "ref": "Model!D4"}, "- Cite as: [northwind.xlsx]Model!D4"),
            "trace": ({"path": "northwind.xlsx", "ref": "Summary!B4"}, "Growth → Inputs!B4"),
        }
        for tool, (arguments, expected) in cases.items():
            with self.subTest(tool=tool):
                self.assertIn(expected, await self.call(tool, **arguments))

    async def test_paths_outside_the_folder_are_refused(self):
        for path in ("../private.xlsx", str(self.base / "private.xlsx")):
            with self.subTest(path=path), self.assertRaisesRegex(ToolError, "outside the folder"):
                await self.call("outline", path=path)

    async def test_problems_reach_the_model_as_messages(self):
        with self.assertRaisesRegex(ToolError, "Did you mean Model"):
            await self.call("read", path="northwind.xlsx", ref="Modle!A1")
        with self.assertRaisesRegex(ToolError, "No such workbook"):
            await self.call("outline", path="missing.xlsx")

    async def test_a_changed_file_is_read_again(self):
        before = await self.call("outline", path="northwind.xlsx")
        replacement = Workbook()
        replacement.active.title = "Fresh"
        replacement.save(self.root / "northwind.xlsx")
        after = await self.call("outline", path="northwind.xlsx")
        self.assertNotIn("## Fresh", before)
        self.assertIn("## Fresh", after)

    async def test_workbooks_are_kept_within_a_budget_of_cells(self):
        for name in ("a.xlsx", "b.xlsx", "c.xlsx"):
            shutil.copy(self.root / "northwind.xlsx", self.root / name)
        cells = sum(len(sheet.cells) for sheet in load(self.root / "a.xlsx").sheets.values())
        with patch("sheetscope.server.MAX_KEPT_CELLS", 2 * cells), patch("sheetscope.server.load", wraps=load) as reads:
            for name in ("a.xlsx", "b.xlsx", "c.xlsx", "c.xlsx", "b.xlsx", "a.xlsx"):
                await self.call("outline", path=name)
        self.assertEqual([call.args[0].name for call in reads.call_args_list], ["a.xlsx", "b.xlsx", "c.xlsx", "a.xlsx"])

    async def test_calls_that_arrive_together_read_a_workbook_once(self):
        def slow(path):
            time.sleep(0.3)
            return load(path)

        with patch("sheetscope.server.load", side_effect=slow) as reads:
            texts = await asyncio.gather(
                self.call("outline", path="northwind.xlsx"),
                self.call("find", path="northwind.xlsx", query="1,597"),
                self.call("cell", path="northwind.xlsx", ref="Model!D4"),
            )
        self.assertEqual(reads.call_count, 1)
        self.assertIn("Model!D4 shows 1,452", texts[2])

    async def test_a_read_that_fails_is_shared_and_tried_again_later(self):
        (self.root / "bad.xlsx").write_bytes(b"PK not a workbook")

        def slow(path):
            time.sleep(0.3)
            return load(path)

        with patch("sheetscope.server.load", side_effect=slow) as reads:
            results = await asyncio.gather(
                *(self.call("outline", path="bad.xlsx") for _ in range(2)), return_exceptions=True
            )
        self.assertEqual(reads.call_count, 1)
        self.assertTrue(all(isinstance(r, ToolError) and "damaged" in str(r) for r in results), results)
        northwind.build(self.root / "bad.xlsx")
        self.assertIn("## Model", await self.call("outline", path="bad.xlsx"))

    async def test_lock_files_hidden_files_and_links_out_are_not_listed(self):
        (self.root / "~$northwind.xlsx").write_bytes(b"lock")
        (self.root / "._northwind.xlsx").write_bytes(b"resource fork")
        (self.root / "gone.xlsx").symlink_to(self.base / "nowhere.xlsx")
        (self.root / "outside.xlsx").symlink_to(self.base / "private.xlsx")
        for folder, name in ((".cache", "copy.xlsx"), ("q3", "model.xlsm")):
            (self.root / folder).mkdir()
            shutil.copy(self.root / "northwind.xlsx", self.root / folder / name)
        listed = [line.split(" (")[0] for line in (await self.call("workbooks")).splitlines()]
        self.assertEqual(listed, ["- northwind.xlsx", "- q3/model.xlsm"])

    async def test_bad_paths_get_a_plain_message_without_the_folder(self):
        (self.root / "folder.xlsx").mkdir()
        for path in ("bad\x00name.xlsx", "folder.xlsx"):
            with self.subTest(path=path), self.assertRaises(ToolError) as caught:
                await self.call("outline", path=path)
            self.assertNotIn(str(self.root), str(caught.exception))

    async def test_the_command_serves_over_stdio(self):
        command = StdioServerParameters(
            command=sys.executable, args=["-m", "sheetscope", "mcp", "--root", str(self.root)]
        )
        with tempfile.TemporaryFile("w+") as log:
            async with stdio_client(command, errlog=log) as (read, write), ClientSession(read, write) as session:
                started = await session.initialize()
                self.assertEqual(started.server_info.name, "sheetscope")
                result = await session.call_tool("cell", {"path": "northwind.xlsx", "ref": "Model!D4"})
                self.assertFalse(result.is_error)
                self.assertIn("Model!D4 shows 1,452", result.content[0].text)
                refused = await session.call_tool("read", {"path": "../private.xlsx", "ref": "Model"})
                self.assertTrue(refused.is_error)
