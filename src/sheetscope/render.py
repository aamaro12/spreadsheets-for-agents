"""Markdown views of each step's result, compact enough for a model's context window."""

from .refs import quote


def markdown(result: dict) -> str:
    """Render a dict returned by one of Book's steps."""
    return _VIEWS[result["kind"]](result)


def _title(result, extra=""):
    return f"**{result['workbook']}** · sha256 {result['sha256'][:12]}" + (f" · {extra}" if extra else "")


def _text(value, limit=160):
    text = str(value)
    if len(text) > limit:
        text = text[: limit - 1] + "…"
    return text.replace("|", "\\|").replace("\r\n", "\n").replace("\n", "<br>")


def _count(number, word):
    return f"{number:,} {word}{'' if number == 1 else 's'}"


def _shown(item):
    return "(no saved value)" if item.get("saved") is False else _text(item.get("display", ""))


def _outline(result):
    lines = [_title(result, _count(len(result["sheets"]), "sheet"))]
    if result["cite_as"]:
        lines.append(f"Cite cells like {result['cite_as']}.")
    for sheet in result["sheets"]:
        facts = [sheet["range"] or "empty", _count(sheet["cells"], "cell")]
        if sheet["state"] != "visible":
            facts.insert(0, "very hidden" if sheet["state"] == "veryHidden" else sheet["state"])
        if sheet["formulas"]:
            facts.append(_count(sheet["formulas"], "formula"))
        if sheet.get("unsaved"):
            facts.append(f"{sheet['unsaved']:,} without saved values")
        if sheet.get("errors"):
            facts.append(_count(sheet["errors"], "error"))
        lines += ["", f"## {sheet['name']} · " + " · ".join(facts)]
        for key, words in (("hidden_rows", "Hidden rows"), ("hidden_columns", "Hidden columns")):
            if sheet.get(key):
                lines.append(f"{words}: {sheet[key]}")
        if sheet.get("comments"):
            lines.append(f"Cell comments: {sheet['comments']}")
        for block in sheet["blocks"]:
            line = f"- {block['range']}"
            if "title" in block:
                line += f" · {_text(block['title'])}"
            if "columns" in block:
                heads = " · ".join(f"{c['col']} {_text(c['text'], 40)}" for c in block["columns"])
                more = f" · {block['more_columns']} more" if block.get("more_columns") else ""
                line += f" · header row {block['header_row']}: {heads}{more}"
            lines.append(line)
            if block["labels"]:
                rows = " · ".join(f"{label['row']} {_text(label['text'], 40)}" for label in block["labels"])
                more = f" · {block['more_labels']} more" if block.get("more_labels") else ""
                lines.append(f"  rows: {rows}{more}")
        if sheet.get("more_blocks"):
            lines.append(f"- {sheet['more_blocks']} more blocks")
    if result["names"]:
        names = []
        for name in result["names"]:
            text = f"{name['name']} = {name['refers_to']}"
            if "value" in name:
                text += f" ({_text(name['value'], 40)})"
            if name.get("scope"):
                text += f" on {name['scope']} only"
            if name.get("hidden"):
                text += " (hidden)"
            names.append(text)
        more = f" · {result['more_names']} more" if result.get("more_names") else ""
        lines += ["", "Names: " + " · ".join(names) + more]
    if result.get("hidden_names"):
        lines += ["", f"Hidden names, not listed: {result['hidden_names']:,}"]
    if result["tables"]:
        tables = [
            f"{t['name']} = {quote(t['sheet'])}!{t['range']} ({', '.join(t['columns'])})" for t in result["tables"]
        ]
        lines += ["", "Tables: " + " · ".join(tables)]
    lines += [f"\nWarning: {warning}" for warning in result["warnings"]]
    return "\n".join(lines)


def _find(result):
    matches, total, query = result["matches"], result["total"], result["query"]
    where = " in formulas" if result["searched"] == "formulas" else ""
    if not total:
        return f'{_title(result)}\nNo cell matches "{query}"{where}.'
    count = f"{total:,} cells match" if total > 1 else "1 cell matches"
    shown = f"; the first {len(matches)} are shown" if len(matches) < total else ""
    lines = [_title(result), f'{count} "{query}"{where}{shown}.', ""]
    with_formulas = any("formula" in m for m in matches)
    lines.append("| Cell | Shows | Row | Column |" + (" Formula |" if with_formulas else ""))
    lines.append("|---|---|---|---|" + ("---|" if with_formulas else ""))
    for m in matches:
        ref = _text(f"{quote(m['sheet'])}!{m['cell']}") + (f" (hidden {m['hidden']})" if m.get("hidden") else "")
        shown = _shown(m) + (f" (in {m['units']})" if m.get("units") else "")
        cells = [ref, shown, _text(m.get("label", ""), 60), _text(m.get("header", ""), 60)]
        if with_formulas:
            cells.append(_text(m.get("formula", "")))
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _read(result):
    where = _text(f"{quote(result['sheet'])}!{result['range']}") + (" (hidden sheet)" if result.get("hidden") else "")
    lines = [_title(result, where)]
    if not result["cells"]:
        return lines[0] + "\nThe range is empty."
    by_place = {item["cell"]: _shown(item) for item in result["cells"]}
    head = ["|   |"]
    for column in result["columns"]:
        name = column["col"] + (f" · {_text(column['header'], 40)}" if "header" in column else "")
        head.append(f" {name}{' (hidden)' if column.get('hidden') else ''} |")
    lines += ["", "".join(head), "|---|" + "---|" * len(result["columns"])]
    for row in result["rows"]:
        name = str(row["row"]) + (f" · {_text(row['label'], 40)}" if "label" in row else "")
        name += " (hidden)" if row.get("hidden") else ""
        values = [by_place.get(f"{column['col']}{row['row']}", "") for column in result["columns"]]
        lines.append(f"| {name} | " + " | ".join(values) + " |")
    if result["truncated"]:
        shown = f"{_count(len(result['rows']), 'row')} and {_count(len(result['columns']), 'column')}"
        lines.append(f"\nOnly the first {shown} are shown; read a narrower range for the rest.")
    if result["formulas"]:
        lines += ["", "Formulas, each written for the first cell of its range:"]
        for group in result["formulas"]:
            kind = " (one array formula)" if group.get("array") else ""
            lines.append(f"- {group['range']}{kind}: {_text(group['formula'], 400)}")
    if result["merged"]:
        lines.append("\nMerged: " + ", ".join(result["merged"]))
    return "\n".join(lines)


def _cell(result):
    ref = _text(f"{quote(result['sheet'])}!{result['cell']}")
    lines = [_title(result), f"{ref} shows {_shown(result) or '(empty)'}", ""]
    if result["type"] != "empty":
        lines.append(f"- Value: {_text(result['value'], 400)} ({result['type']}, format {result['number_format']})")
    if "formula" in result:
        array = f" (one array formula over {result['array']})" if "array" in result else ""
        unsaved = " (no saved value)" if result.get("saved") is False else ""
        lines.append(f"- Formula: {result['formula']}{array}{unsaved}")
    context = [
        f"{word}: {_text(result[key], 60)}" for key, word in (("label", "Row"), ("header", "Column")) if key in result
    ]
    if context:
        lines.append("- " + " · ".join(context))
    for key, word in (("hidden", "Hidden"), ("merged", "Merged"), ("comment", "Comment"), ("link", "Link")):
        if key in result:
            lines.append(f"- {word}: {_text(result[key], 400)}")
    if "precedents" in result:
        reads = "; ".join(_reference(item) for item in result["precedents"]) or "no cells"
        lines.append(f"- Reads: {reads}")
    lines += [f"- Note: {note}" for note in result.get("notes", [])]
    users = ", ".join(result["dependents"]) or "no formulas"
    more = f" and {result['dependents_total'] - len(result['dependents'])} more" if "dependents_total" in result else ""
    lines.append(f"- Used by: {users}{more}")
    lines.append(f"- Cite as: {result['citation']}")
    return "\n".join(lines)


def _reference(item):
    """One cell or range read by a formula, with what it shows."""
    name = f"{item['via']} → {item['ref']}" if "via" in item and item["via"] != item["ref"] else item["ref"]
    if "cells" in item:
        values = f": {', '.join(_text(v, 40) for v in item['values'])}" if "values" in item else ""
        return f"{name} ({_count(item['cells'], 'cell')}, {_count(item['formulas'], 'formula')}){values}"
    context = ", ".join(_text(item[key], 40) for key in ("label", "header") if key in item)
    return f"{name}{f' ({context})' if context else ''} = {_shown(item)}"


def _trace(result):
    direction = "the cells it reads" if result["direction"] == "precedents" else "the formulas that use it"
    lines = [_title(result), f"{result['root']['ref']} and {direction}:", ""]

    def walk(node, depth):
        if "note" in node:
            lines.append("  " * depth + f"- Note: {node['note']}")
            return
        line = "  " * depth + "- " + _reference(node)
        if "formula" in node:
            line += f" · {_text(node['formula'], 200)}"
        for key, words in (("repeat", "see above"), ("circular", "circular"), ("more", "more beyond this depth")):
            if node.get(key):
                line += f" ({words})"
        if node.get("skipped"):
            line += f" (the first {_count(len(node['children']), 'formula')} followed, {node['skipped']:,} more not)"
        lines.append(line)
        for child in node.get("children", []):
            walk(child, depth + 1)

    walk(result["root"], 0)
    return "\n".join(lines)


_VIEWS = {"outline": _outline, "find": _find, "read": _read, "cell": _cell, "trace": _trace}
