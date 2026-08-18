"""
xlsx_read — bounded-memory READ side of the surgeon.

Everything here streams or reads only small parts; nothing loads a large
sheet whole. Shared strings are resolved SELECTIVELY: the callers collect
the si indices they actually need and a single streaming pass over
xl/sharedStrings.xml captures just those — the full table on this workbook
would be enormous (the cumulative Raw tab alone has ~480k text rows).
"""

import codecs
import re
import zipfile

CHUNK = 4 * 1024 * 1024
_CELL_RE = re.compile(
    r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>|<c r="([A-Z]+)(\d+)"([^>]*)/>', re.S)
_V_RE = re.compile(r"<v>(.*?)</v>", re.S)
_F_RE = re.compile(r"<f[^>]*>(.*?)</f>", re.S)
_IS_RE = re.compile(r"<is>.*?</is>", re.S)
_T_RE = re.compile(r"<t[^>]*>(.*?)</t>", re.S)


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _unescape(s: str) -> str:
    for ent, ch in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                    ("&apos;", "'"), ("&amp;", "&")):
        s = s.replace(ent, ch)
    return s


def parse_row_cells(row_xml: str) -> dict:
    """{'A': {'t': type, 'v': raw <v>, 'f': formula, 'is': inline text}}"""
    out = {}
    for m in _CELL_RE.finditer(row_xml):
        if m.group(1) is not None:
            col, attrs, body = m.group(1), m.group(3) or "", m.group(4) or ""
        else:
            col, attrs, body = m.group(5), m.group(7) or "", ""
        tm = re.search(r'\bt="([^"]+)"', attrs)
        cell = {"t": tm.group(1) if tm else None, "v": None, "f": None, "is": None}
        vm = _V_RE.search(body)
        if vm:
            cell["v"] = _unescape(vm.group(1))
        fm = _F_RE.search(body)
        if fm:
            cell["f"] = "=" + _unescape(re.sub(r"\s+", " ", fm.group(1)).strip())
        im = _IS_RE.search(body)
        if im:
            cell["is"] = _unescape("".join(_T_RE.findall(im.group(0))))
        out[col] = cell
    return out


def stream_rows(zf: zipfile.ZipFile, part: str, first_row: int, last_row: int,
                scan_cap: int = 96 * 1024 * 1024) -> dict:
    """{row_num: {col: cell}} for rows first..last, streaming the part from
    the top and stopping as soon as the range is passed. Safe on any part
    size when the target rows sit near the head (headers, dashboards)."""
    rows: dict = {}
    dec = codecs.getincrementaldecoder("utf-8")("replace")
    buf, scanned = "", 0
    row_re = re.compile(r'<row r="(\d+)"(?:\s[^>]*)?(?:/>|>.*?</row>)', re.S)
    with zf.open(part) as f:
        while scanned < scan_cap:
            b = f.read(CHUNK)
            scanned += len(b) if b else 0
            buf += dec.decode(b, final=not b)
            pos = 0
            for m in row_re.finditer(buf):
                rn = int(m.group(1))
                pos = m.end()
                if rn > last_row:
                    return rows
                if rn >= first_row:
                    rows[rn] = parse_row_cells(m.group(0))
            buf = buf[max(pos, len(buf) - 65536):] if pos else buf[-65536:]
            if not b:
                break
    return rows


def resolve_shared(zf: zipfile.ZipFile, needed: set) -> dict:
    """Stream xl/sharedStrings.xml capturing ONLY the needed indices."""
    out: dict = {}
    if not needed or "xl/sharedStrings.xml" not in zf.namelist():
        return out
    want = set(int(i) for i in needed)
    hi = max(want)
    dec = codecs.getincrementaldecoder("utf-8")("replace")
    buf, idx = "", 0
    si_re = re.compile(r"<si>(.*?)</si>", re.S)
    with zf.open("xl/sharedStrings.xml") as f:
        while True:
            b = f.read(CHUNK)
            buf += dec.decode(b, final=not b)
            pos = 0
            for m in si_re.finditer(buf):
                if idx in want:
                    out[idx] = _unescape("".join(_T_RE.findall(m.group(1))))
                    want.discard(idx)
                idx += 1
                pos = m.end()
                if idx > hi or not want:
                    return out
            buf = buf[pos:] if pos else buf
            if not b:
                return out


def cell_value(cell: dict, shared: dict):
    """Best-effort typed value; formula cells yield their cached value."""
    if cell is None:
        return None
    t = cell.get("t")
    if t == "s":
        try:
            return shared.get(int(cell.get("v")))
        except (TypeError, ValueError):
            return None
    if t == "inlineStr":
        return cell.get("is")
    if t == "str":
        return cell.get("v")
    if t == "b":
        return cell.get("v") == "1"
    v = cell.get("v")
    if v in (None, ""):
        return None
    try:
        fv = float(v)
        return int(fv) if fv.is_integer() and "e" not in v.lower() else fv
    except ValueError:
        return v


def shared_indices(cells_iter) -> set:
    need = set()
    for cell in cells_iter:
        if cell and cell.get("t") == "s":
            try:
                need.add(int(cell.get("v")))
            except (TypeError, ValueError):
                pass
    return need


def part_dimension(zf: zipfile.ZipFile, part: str) -> str | None:
    with zf.open(part) as f:
        head = f.read(8192).decode("utf-8", "replace")
    m = re.search(r'<dimension ref="([^"]+)"', head)
    return m.group(1) if m else None


def parse_range(rng: str):
    """'W6:W24' -> (col_lo, col_hi, row_lo, row_hi); 'B:E' -> rows None."""
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", rng.strip().upper())
    if m:
        return (col_index(m.group(1)), col_index(m.group(3)),
                int(m.group(2)), int(m.group(4)))
    m = re.match(r"^([A-Z]+):([A-Z]+)$", rng.strip().upper())
    if m:
        return col_index(m.group(1)), col_index(m.group(2)), None, None
    m = re.match(r"^([A-Z]+)(\d+)$", rng.strip().upper())
    if m:
        c = col_index(m.group(1))
        return c, c, int(m.group(2)), int(m.group(2))
    raise ValueError(f"unsupported range {rng!r}")
