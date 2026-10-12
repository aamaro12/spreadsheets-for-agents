"""Excel number formats: the text a cell shows on screen, so 0.125 formatted 0.0% shows 12.5%."""

from datetime import date, datetime, time, timedelta
from decimal import ROUND_HALF_UP, Decimal, localcontext
from functools import lru_cache
import math
import re

from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import from_excel

_PIECES = re.compile(r'"[^"]*"?|\\.?|\[[^\]]*\]?|.', re.S)
_BRACKET = re.compile(r'("[^"]*"?|\\.)|\[([^\]]*)\]')
_CONDITION = re.compile(r"(<=|>=|<>|<|>|=)\s*(-?[0-9.]+)")
_NUMBER = re.compile(r'"([^"]*)"?|\\(.)|_.|\*.|(General|@)|([0#?.,%])|([Ee][+-])|(.)', re.S | re.I)
_DATE = re.compile(r'"([^"]*)"?|\\(.)|_.|\*.|(\[[hms]+\]|y+|m+|d+|h+|s+|AM/PM|A/P|\.0+)|(.)', re.S | re.I)
_FRACTION = re.compile(r"[0#?]\s*/\s*[0-9#?]")
_QUOTED = re.compile(r'"[^"]*"?|\\.')
_TEXT = re.compile(r'"([^"]*)"?|\\(.)|_.|\*.|(@)|(.)', re.S)
_COMPARE = {
    "<": float.__lt__,
    ">": float.__gt__,
    "=": float.__eq__,
    "<=": float.__le__,
    ">=": float.__ge__,
    "<>": float.__ne__,
}
_MONTHS = "January February March April May June July August September October November December".split()
_DAYS = "Monday Tuesday Wednesday Thursday Friday Saturday Sunday".split()


def display(value, fmt="General") -> str:
    """The text Excel shows for a value in a given number format."""
    fmt = fmt or "General"
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time)):
        return _date(value, fmt)
    if isinstance(value, timedelta):
        return _duration(value, fmt)
    if isinstance(value, (int, float)):
        if _dated(fmt):
            return _date(from_excel(value), fmt)
        return _number(value, fmt)
    return _text(str(value), fmt)


_dated = lru_cache(maxsize=1024)(is_date_format)
_SILENT = re.compile(r"General|\[(?!\$)[^\]]*\]", re.I)  # colours, conditions and elapsed times show no letters


def may_show(text: str, fmt: str) -> bool:
    """Whether a number in this format could show this text: any letters in it must come from the format."""
    return all(letter in _letters(fmt) for letter in text if letter.isalpha())


@lru_cache(maxsize=1024)
def _letters(fmt):
    return frozenset(c for c in _SILENT.sub("", fmt or "General").casefold() if c.isalpha()) | {"e"}  # e: 1E+20


@lru_cache(maxsize=1024)
def _sections(fmt):
    """Split a format at its section separators, skipping quoted text, escapes and brackets."""
    sections = [""]
    for piece in _PIECES.findall(fmt):
        if piece == ";":
            sections.append("")
        else:
            sections[-1] += piece
    return tuple(sections)


@lru_cache(maxsize=2048)
def _clean(section):
    """Drop colours and locale codes, keep currency symbols, and return the section with its condition."""
    condition = None

    def bracket(match):
        nonlocal condition
        if match[1]:
            return match[1]
        body = match[2]
        if found := _CONDITION.fullmatch(body):
            condition = (found[1], float(found[2]))
            return ""
        if body.startswith("$"):
            symbol = body[1:].split("-")[0]
            return f'"{symbol}"' if symbol else ""
        if re.fullmatch(r"h+|m+|s+", body, re.I):
            return f"[{body}]"
        return ""

    return _BRACKET.sub(bracket, section), condition


def _number(value, fmt):
    sections = _cleaned(fmt)
    if any(condition for _, condition in sections[:2]):
        chosen = next((s for s, c in sections[:2] if c and _COMPARE[c[0]](float(value), c[1])), None)
        if chosen is None:
            chosen = next((s for s, c in sections[1:] if not c), sections[-1][0])
        return ("-" if value < 0 else "") + _layout(abs(value), chosen)
    texts = [s for s, _ in sections]
    if len(texts) >= 3 and value == 0:
        return _layout(0, texts[2])
    if len(texts) >= 2 and value < 0:
        return _layout(-value, texts[1])
    return ("-" if value < 0 else "") + _layout(abs(value), texts[0])


@lru_cache(maxsize=1024)
def _cleaned(fmt):
    return tuple(_clean(section) for section in _sections(fmt))


def _layout(value, section):
    """Place a non-negative number into one format section."""
    kind, before, after, percent, spec = _plan(section)
    if kind == "text":
        return before
    value *= 100**percent
    if kind == "general":
        return before + _general(value) + after
    if kind == "scientific":
        return before + _scientific(value, *spec) + after
    whole, fraction, dot, scale = spec
    return before + _fixed(value / 1000**scale, whole, fraction, dot) + after


@lru_cache(maxsize=4096)
def _plan(section):
    """How a format section places a number: the text before and after it, its percent scaling and digits."""
    parts, percent = [], 0
    for quoted, escaped, general, holder, exponent, other in _NUMBER.findall(section):
        if holder == "%":
            percent += 1
            parts.append(("text", "%"))
        elif holder:
            parts.append(("digit", holder))
        elif general:
            parts.append(("general", ""))
        elif exponent:
            parts.append(("exponent", exponent))
        elif quoted or escaped or other:
            parts.append(("text", quoted or escaped or other))
    slots = [i for i, (kind, _) in enumerate(parts) if kind != "text"]
    if not slots:
        return "text", "".join(text for _, text in parts), "", 0, None
    if _FRACTION.search(_QUOTED.sub("", section)):
        parts, slots = [("general", "")], [0]
    first, last = slots[0], slots[-1]
    before = "".join(text for kind, text in parts[:first] if kind == "text")
    after = "".join(text for kind, text in parts[last + 1 :] if kind == "text")
    middle = parts[first : last + 1]
    if any(kind == "general" for kind, _ in middle):
        return "general", before, after, percent, None
    spec = "".join(text if kind == "digit" else "E" + text[1] for kind, text in middle if kind != "text")
    if "E" in spec:
        mantissa, _, power = spec.partition("E")
        return "scientific", before, after, percent, (mantissa, power[0], power[1:])
    whole, dot, fraction = spec.partition(".")
    scale = len(whole) - len(whole.rstrip(",")) + len(fraction) - len(fraction.rstrip(","))
    return "fixed", before, after, percent, (whole.rstrip(","), fraction.replace(",", ""), bool(dot), scale)


def _round(value, places):
    with localcontext() as context:
        context.prec = 400
        return Decimal(f"{value:.15g}").quantize(Decimal(1).scaleb(-places), ROUND_HALF_UP)


def _fixed(value, whole_spec, fraction_spec, dot):
    places = sum(ch in "0#?" for ch in fraction_spec)
    whole, _, fraction = f"{_round(value, places):f}".partition(".")
    fraction = fraction.ljust(places, "0")[:places]
    keep = fraction_spec.count("0")
    while len(fraction) > keep and fraction.endswith("0"):
        fraction = fraction[:-1]
    number, minimum = int(whole), whole_spec.count("0")
    if number == 0 and minimum == 0:
        whole = ""
    else:
        whole = (f"{number:,}" if "," in whole_spec else str(number)).zfill(minimum)
    return whole + ("." + fraction if dot else "")


def _scientific(value, mantissa_spec, sign, power_spec):
    whole, dot, fraction = mantissa_spec.partition(".")
    places = sum(ch in "0#?" for ch in fraction)
    power = 0 if value == 0 else math.floor(math.log10(value))
    if value and _round(value / 10**power, places) >= 10:
        power += 1
    mantissa = _fixed(value / 10**power if value else 0, whole or "0", fraction, bool(dot))
    shown_sign = "-" if power < 0 else ("+" if sign == "+" else "")
    return f"{mantissa}E{shown_sign}{str(abs(power)).zfill(power_spec.count('0'))}"


def _general(value):
    if isinstance(value, int) or (value.is_integer() and abs(value) < 1e15):
        return str(int(value))
    text = f"{value:.10g}"
    if "e" in text:
        mantissa, power = text.split("e")
        return f"{mantissa}E{int(power):+03d}"
    return text


def _text(value, fmt):
    pieces = _text_plan(fmt)
    return value if pieces is None else "".join(value if piece is None else piece for piece in pieces)


@lru_cache(maxsize=1024)
def _text_plan(fmt):
    """The pieces a format shows around text, with None where the text goes; None when it shows text as is."""
    sections = _sections(fmt)
    if len(sections) >= 4:
        section = sections[3]
    elif len(sections) == 1 and "@" in _TEXT.sub(lambda m: m[3] or "", sections[0]):
        section = sections[0]
    else:
        return None
    return tuple(None if at else quoted or escaped or other for quoted, escaped, at, other in _TEXT.findall(section))


def _date(value, fmt):
    plan, twelve = _date_plan(fmt)
    if not any(piece[0] == "code" for piece in plan):
        return _iso(value)
    if isinstance(value, time):
        value = datetime.combine(date(1900, 1, 1), value)
    elif not isinstance(value, datetime):
        value = datetime.combine(value, time())
    return "".join(piece[1] if piece[0] == "text" else _date_part(value, *piece[1:], twelve) for piece in plan)


@lru_cache(maxsize=1024)
def _date_plan(fmt):
    """A date format as pieces of text and codes, each code with what it stands for: m is a month or a minute."""
    section = _clean(_sections(fmt)[0])[0]
    plan = [
        ("code", code) if code else ("text", quoted or escaped or other)
        for quoted, escaped, code, other in _DATE.findall(section)
    ]
    codes = [i for i, (kind, _) in enumerate(plan) if kind == "code"]
    originals = [plan[i][1].strip("[]") for i in codes]
    twelve = any(code.upper() in ("AM/PM", "A/P") for code in originals)
    for position, (i, code) in enumerate(zip(codes, originals)):
        lower = code.lower()
        if lower[0] == "m" and len(lower) <= 2:
            previous = originals[position - 1].lower() if position else ""
            following = originals[position + 1].lower() if position + 1 < len(codes) else ""
            if previous.startswith("h") or following.startswith("s"):
                lower = "minute" + lower
        plan[i] = ("code", lower, code)
    return tuple(plan), twelve


def _date_part(value, code, original, twelve):
    hour = value.hour % 12 or 12 if twelve else value.hour
    if code[0] == "y":
        return f"{value.year:04d}" if len(code) > 2 else f"{value.year % 100:02d}"
    if code.startswith("minute"):
        return f"{value.minute:02d}" if code == "minutemm" else str(value.minute)
    if code[0] == "m":
        name = _MONTHS[value.month - 1]
        return {1: str(value.month), 2: f"{value.month:02d}", 3: name[:3], 4: name}.get(len(code), name[0])
    if code[0] == "d":
        name = _DAYS[value.weekday()]
        return {1: str(value.day), 2: f"{value.day:02d}", 3: name[:3]}.get(len(code), name)
    if code[0] == "h":
        return f"{hour:02d}" if len(code) > 1 else str(hour)
    if code[0] == "s":
        return f"{value.second:02d}" if len(code) > 1 else str(value.second)
    if code[0] == ".":
        return "." + f"{value.microsecond:06d}"[: len(code) - 1]
    morning = value.hour < 12
    if code == "a/p":
        text = "A" if morning else "P"
    else:
        text = "AM" if morning else "PM"
    return text if original[0].isupper() else text.lower()


def _duration(value, fmt):
    """A span of time, such as 26:30 for 26 hours and 30 minutes formatted [h]:mm."""
    total = round(value.total_seconds())
    sign, total = ("-" if total < 0 else ""), abs(total)
    out = []
    for quoted, escaped, code, other in _DATE.findall(_clean(_sections(fmt)[0])[0]):
        unit = code.strip("[]").lower()[:1]
        if unit not in ("h", "m", "s", "d"):
            out.append(code if code and not code.startswith(".") else quoted or escaped or other)
            continue
        elapsed = code.startswith("[")
        number = {
            "d": total // 86400,
            "h": total // 3600 if elapsed else total // 3600 % 24,
            "m": total // 60 if elapsed else total // 60 % 60,
            "s": total if elapsed else total % 60,
        }[unit]
        out.append(str(number).zfill(min(len(code.strip("[]")), 2)))
    return sign + "".join(out) if any(code for _, _, code, _ in _DATE.findall(fmt)) else sign + _clock(total)


def _clock(seconds):
    return f"{seconds // 3600}:{seconds // 60 % 60:02d}:{seconds % 60:02d}"


def _iso(value):
    if isinstance(value, datetime):
        if value.time() == time():
            return value.date().isoformat()
        return value.isoformat(sep=" ", timespec="seconds")
    return value.isoformat()
