"""
xlsx_surgeon — direct ZIP/XML surgery on large .xlsx files that the Graph
workbook API cannot open.

Design constraints (why this is corruption-safe):
  * Appended/updated cells use INLINE strings (t="inlineStr") -> sharedStrings
    is never touched, so no existing string index can be invalidated.
  * New cells use style index 0 or a caller-supplied existing style index ->
    styles.xml is never touched.
  * calcChain.xml is deleted (Excel rebuilds it) and <calcPr fullCalcOnLoad="1"/>
    is set -> Excel fully recalculates on open, evaluating every formula we
    wrote and every formula that depends on the data we appended.
  * Every part we do not explicitly modify is streamed through byte-identical.
  * Large sheet parts are processed on disk in chunks (bounded memory).

Operations:
  set_cells(sheet, {"B2": "June'26", "U5": 123.4, "V4": "=SUM(A:A)"})
  append_rows(sheet, [[...], [...]])              # values / formulas ("=..")
  add_sheet(name, rows)                           # new worksheet at the end
"""

import os
import re
import shutil
import zipfile
import tempfile

CHUNK = 4 * 1024 * 1024
XLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

try:
    import resource   # Linux/Render only — absent on Windows, diagnostic-only

    def _mem_log(label: str) -> None:
        mb = round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 1)
        print(f"[mem] peak RSS {label}: {mb} MB", flush=True)
except ImportError:
    def _mem_log(label: str) -> None:
        pass


# ---------------------------------------------------------------- helpers
def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def split_ref(ref: str):
    m = re.match(r"^([A-Za-z]+)(\d+)$", ref)
    if not m:
        raise ValueError(f"bad cell ref {ref!r}")
    return m.group(1).upper(), int(m.group(2))


def esc(s: str) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def cell_xml(ref: str, value, style: int = 0) -> str:
    """Serialize one cell. Strings are inline (never sharedStrings)."""
    s = f' s="{style}"' if style else ""
    if value is None or value == "":
        return ""
    if isinstance(value, str) and value.startswith("="):
        return f'<c r="{ref}"{s}><f>{esc(value[1:])}</f></c>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{s} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s}><v>{value!r}</v></c>'
    return (f'<c r="{ref}"{s} t="inlineStr">'
            f'<is><t xml:space="preserve">{esc(value)}</t></is></c>')


def row_xml(row_num: int, values, styles=None) -> str:
    """Builds one <row>. append_rows/add_sheet are the only ops where the
    caller cannot know a row's final number in advance (paste_columns and
    set_cells target explicit refs, so callers substitute {r} themselves) —
    a formula value still containing the literal '{r}' gets it filled in
    here, at the one point row_num is actually known."""
    cells = []
    for i, v in enumerate(values):
        if isinstance(v, str) and "{r}" in v:
            v = v.replace("{r}", str(row_num))
        st = styles[i] if styles and i < len(styles) else 0
        x = cell_xml(f"{col_letter(i + 1)}{row_num}", v, st)
        if x:
            cells.append(x)
    return f'<row r="{row_num}">{"".join(cells)}</row>'


# ---------------------------------------------------------------- surgeon
class XlsxSurgeon:
    def __init__(self, src_path: str, workdir: str | None = None):
        self.src = src_path
        self.workdir = workdir or tempfile.mkdtemp(prefix="surgeon_")
        os.makedirs(self.workdir, exist_ok=True)
        self._ops = []               # (kind, sheet_or_name, payload)
        with zipfile.ZipFile(self.src) as zf:
            self._names = zf.namelist()
            self._wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
            self._wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            self._ctypes = zf.read("[Content_Types].xml").decode("utf-8")
        self._sheet_parts = self._map_sheets()

    # -- workbook topology -------------------------------------------------
    def _map_sheets(self):
        rid_to_target = {}
        for m in re.finditer(r'<Relationship\b[^>]*/?>', self._wb_rels):
            tag = m.group(0)
            rid = re.search(r'Id="([^"]+)"', tag)
            tgt = re.search(r'Target="([^"]+)"', tag)
            typ = re.search(r'Type="[^"]*/(\w+)"', tag)
            if rid and tgt and typ and typ.group(1) == "worksheet":
                target = tgt.group(1)
                if not target.startswith("/"):
                    target = "xl/" + target
                else:
                    target = target.lstrip("/")
                rid_to_target[rid.group(1)] = target
        sheets = {}
        for m in re.finditer(r'<sheet\b[^>]*/?>', self._wb_xml):
            tag = m.group(0)
            name = re.search(r'name="([^"]*)"', tag)
            rid = re.search(r'r:id="([^"]+)"', tag)
            if name and rid and rid.group(1) in rid_to_target:
                sheets[name.group(1)] = rid_to_target[rid.group(1)]
        return sheets

    def sheet_names(self):
        return list(self._sheet_parts)

    # -- public ops --------------------------------------------------------
    def set_cells(self, sheet: str, cells: dict, style_from: dict | None = None):
        # may target a sheet created by duplicate_sheet earlier in the same job
        pending_dup = any(o[0] == "dup" and o[1] == sheet for o in self._ops)
        if sheet not in self._sheet_parts and not pending_dup:
            raise KeyError(f"sheet {sheet!r} not found; have {self.sheet_names()}")
        self._ops.append(("set", sheet, (cells, style_from or {})))

    def append_rows(self, sheet: str, rows: list, styles=None):
        if sheet not in self._sheet_parts:
            raise KeyError(f"sheet {sheet!r} not found; have {self.sheet_names()}")
        self._ops.append(("append", sheet, (rows, styles)))

    def add_sheet(self, name: str, rows: list, styles=None):
        if name in self._sheet_parts:
            raise KeyError(f"sheet {name!r} already exists")
        self._ops.append(("add", name, (rows, styles)))

    def duplicate_sheet(self, source: str, new_name: str):
        """Byte-copy an existing sheet (formulas, styles refs intact) under a
        new name. Drawing/table references are stripped from the copy."""
        if source not in self._sheet_parts:
            raise KeyError(f"sheet {source!r} not found; have {self.sheet_names()}")
        if new_name in self._sheet_parts:
            raise KeyError(f"sheet {new_name!r} already exists")
        self._ops.append(("dup", new_name, source))

    def paste_columns(self, sheet: str, anchor: str, rows: list,
                      clear_beyond: bool = True):
        """Overwrite a rectangular block starting at `anchor` (e.g. 'A7') with
        `rows`, preserving all cells outside the block width (formula columns
        to the right survive). Existing rows below the pasted block have the
        block's columns cleared when clear_beyond=True. In-memory: sheet must
        be < 32MB decompressed. May target a sheet created by duplicate_sheet
        in the same job."""
        pending_dup = any(o[0] == "dup" and o[1] == sheet for o in self._ops)
        if sheet not in self._sheet_parts and not pending_dup:
            raise KeyError(f"sheet {sheet!r} not found; have {self.sheet_names()}")
        self._ops.append(("paste", sheet, (anchor, rows, clear_beyond)))

    # -- application -------------------------------------------------------
    def apply(self, dst_path: str):
        # group ops per existing part
        per_part: dict[str, dict] = {}
        new_sheets = []
        dup_sheets = []
        for kind, target, payload in self._ops:
            if kind == "add":
                new_sheets.append((target, payload))
                continue
            if kind == "dup":
                dup_sheets.append((target, payload))   # (new_name, source)
                continue
            if target not in self._sheet_parts:
                continue   # targets a duplicated sheet; handled at duplication time
            part = self._sheet_parts[target]
            per_part.setdefault(part, {"set": {}, "append": [], "paste": []})
            if kind == "set":
                per_part[part]["set"].update(payload[0])
            elif kind == "paste":
                per_part[part]["paste"].append(payload)
            else:
                per_part[part]["append"].append(payload)

        # register new sheets in workbook.xml / rels / content types
        wb_xml, wb_rels, ctypes = self._wb_xml, self._wb_rels, self._ctypes
        new_parts = []
        results: list[dict] = []      # per-target change counts, returned to caller
        # duplicated sheets: read source part, strip relationship refs, queue as new part
        if dup_sheets:
            with zipfile.ZipFile(self.src) as zf:
                for new_name, source in dup_sheets:
                    src_part = self._sheet_parts[source]
                    src_size = zf.getinfo(src_part).file_size
                    print(f"[mem] duplicate_sheet {source!r}: source part is "
                         f"{round(src_size / 1048576, 1)} MB compressed", flush=True)
                    _mem_log(f"before reading {source!r}")
                    xml = zf.read(src_part).decode("utf-8")
                    print(f"[mem] duplicate_sheet {source!r}: decoded to "
                         f"{round(len(xml) / 1_000_000, 1)} MB of text", flush=True)
                    _mem_log(f"after decoding {source!r}")
                    # The copy gets NO _rels part, so ANY surviving r:id is a hard
                    # OPC violation — Excel repairs the file and drops the sheet.
                    # pageSetup carries an r:id whenever the sheet ever had a
                    # print area; hyperlinks/pictures/controls likewise.
                    # Combined into 3 passes (was 20 separate re.sub calls, each
                    # a full-string copy of a real sheet that can be tens of MB)
                    # — a straightforward win regardless of what else is resident.
                    xml = re.sub(
                        r"<(?P<tag>drawing|legacyDrawing|tableParts|pageSetup|"
                        r"hyperlink|picture|oleObject|control)\b[^>]*>.*?</(?P=tag)>",
                        "", xml, flags=re.S)
                    xml = re.sub(
                        r"<(?:drawing|legacyDrawing|tableParts|pageSetup|"
                        r"hyperlink|picture|oleObject|control)\b[^>]*/>", "", xml)
                    xml = re.sub(
                        r"<(?P<wrap>hyperlinks|oleObjects|controls)\s*>\s*</(?P=wrap)>"
                        r"|<(?:hyperlinks|oleObjects|controls)\s*/>", "", xml)
                    _mem_log(f"after stripping relationship tags on {source!r}")
                    self._assert_no_dangling_rids(f"duplicate of {source!r}", xml, ())
                    # duplicated sheet may still need pastes applied to it later:
                    # apply queued paste/set ops that target the NEW name in-memory now.
                    # ONE merged pass, not one per op — see _apply_row_ops.
                    changed = 1                     # the new sheet itself
                    pend = [o for o in self._ops if o[1] == new_name and o[0] in ("paste", "set")]
                    paste_groups = [p[2] for p in pend if p[0] == "paste"]
                    cell_groups = [p[2][0] for p in pend if p[0] == "set"]
                    if paste_groups or cell_groups:
                        _mem_log(f"before paste/set_cells on duplicated {new_name!r}")
                        xml, n = self._apply_row_ops(xml, paste_groups, cell_groups)
                        changed += n
                        _mem_log(f"after paste/set_cells on duplicated {new_name!r}")
                    results.append({"target": new_name, "kind": "duplicate_sheet",
                                    "sourcePartMB": round(src_size / 1048576, 1),
                                    "cellsChanged": changed})
                    new_sheets.append((new_name, ("__RAW__", xml)))
        if new_sheets:
            max_sheet_num = 0
            for p in self._names:
                m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", p)
                if m:
                    max_sheet_num = max(max_sheet_num, int(m.group(1)))
            max_rid = max((int(m.group(1)) for m in
                           re.finditer(r'Id="rId(\d+)"', wb_rels)), default=0)
            max_sheet_id = max((int(m.group(1)) for m in
                                re.finditer(r'sheetId="(\d+)"', wb_xml)), default=0)
            for i, (name, payload_ns) in enumerate(new_sheets, start=1):
                rows, styles = payload_ns
                part = f"xl/worksheets/sheet{max_sheet_num + i}.xml"
                rid = f"rId{max_rid + i}"
                sid = max_sheet_id + i
                wb_xml = wb_xml.replace(
                    "</sheets>",
                    f'<sheet name="{esc(name)}" sheetId="{sid}" r:id="{rid}"/></sheets>')
                wb_rels = wb_rels.replace(
                    "</Relationships>",
                    f'<Relationship Id="{rid}" '
                    f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/worksheet" Target="worksheets/sheet{max_sheet_num + i}.xml"/>'
                    f'</Relationships>')
                ctypes = ctypes.replace(
                    "</Types>",
                    f'<Override PartName="/{part}" ContentType="application/vnd.'
                    f'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
                if rows == "__RAW__":
                    new_parts.append((part, styles))     # styles holds raw xml here
                else:
                    results.append({"target": name, "kind": "add_sheet",
                                    "cellsChanged": max(
                                        1, sum(len(r) for r in rows))})
                    new_parts.append((part, self._build_sheet_xml(rows, styles)))

        # calcChain: drop part + its Content_Types override + workbook rel
        drop = set()
        if "xl/calcChain.xml" in self._names:
            drop.add("xl/calcChain.xml")
            ctypes = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ctypes)
            wb_rels = re.sub(r'<Relationship\b[^>]*calcChain\.xml[^>]*/>', "", wb_rels)

        # force full recalc on open
        if "<calcPr" in wb_xml:
            if "fullCalcOnLoad" in wb_xml:
                wb_xml = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', wb_xml)
            else:
                wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        else:
            wb_xml = wb_xml.replace("</sheets>", '</sheets><calcPr fullCalcOnLoad="1"/>', 1)

        # every r:id the workbook references must have a matching relationship
        self._assert_no_dangling_rids(
            "xl/workbook.xml", wb_xml,
            re.findall(r'Id="([^"]+)"', wb_rels))

        replaced = {
            "xl/workbook.xml": wb_xml.encode("utf-8"),
            "xl/_rels/workbook.xml.rels": wb_rels.encode("utf-8"),
            "[Content_Types].xml": ctypes.encode("utf-8"),
        }

        part_to_name = {v: k for k, v in self._sheet_parts.items()}
        with zipfile.ZipFile(self.src) as zin, \
             zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                if name in drop:
                    continue
                if name in replaced:
                    zout.writestr(name, replaced[name])
                elif name in per_part:
                    tmp_in = os.path.join(self.workdir, "part_in.xml")
                    with zin.open(name) as f, open(tmp_in, "wb") as out:
                        shutil.copyfileobj(f, out, CHUNK)
                    tmp_out = os.path.join(self.workdir, "part_out.xml")
                    changed = self._transform_sheet(
                        tmp_in, tmp_out,
                        per_part[name]["set"],
                        per_part[name]["append"],
                        per_part[name].get("paste", []))
                    # transforms only rewrite <row>/<c> content, which never
                    # carries r:id, and the part's own _rels streams through
                    # untouched — so no dangling-rid scan is needed here.
                    results.append({"target": part_to_name.get(name, name),
                                    "kind": "transform",
                                    "cellsChanged": changed})
                    zi = zipfile.ZipInfo(name)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    with open(tmp_out, "rb") as f:
                        with zout.open(zi, "w", force_zip64=True) as w:
                            shutil.copyfileobj(f, w, CHUNK)
                    os.remove(tmp_in), os.remove(tmp_out)
                else:
                    # untouched: stream through
                    zi = zipfile.ZipInfo(name)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    with zin.open(name) as f:
                        with zout.open(zi, "w", force_zip64=True) as w:
                            shutil.copyfileobj(f, w, CHUNK)
            for part, xml in new_parts:
                # authored in memory with no _rels part: r:ids must be gone
                self._assert_no_dangling_rids(part, xml, ())
                zout.writestr(part, xml)

        if sum(r["cellsChanged"] for r in results) == 0:
            raise ValueError(
                "no op changed anything — refusing to upload an unmodified file")
        return results

    @staticmethod
    def _assert_no_dangling_rids(part_name: str, xml: str, rel_ids) -> None:
        used = set(re.findall(r'r:id="([^"]+)"', xml))
        dangling = used - set(rel_ids)
        if dangling:
            raise ValueError(
                f"{part_name}: relationship ids {sorted(dangling)} have no "
                "matching _rels entry — writing this part would corrupt the file")

    # -- sheet XML construction / transformation ---------------------------
    @staticmethod
    def _build_sheet_xml(rows, styles=None):
        body = "".join(row_xml(i + 1, r, styles) for i, r in enumerate(rows))
        ncols = max((len(r) for r in rows), default=1)
        dim = f"A1:{col_letter(ncols)}{max(len(rows), 1)}"
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<worksheet xmlns="{XLNS}" xmlns:r="http://schemas.openxmlformats.'
                f'org/officeDocument/2006/relationships">'
                f'<dimension ref="{dim}"/><sheetData>{body}</sheetData></worksheet>')

    def _transform_sheet(self, src, dst, set_cells, append_groups, paste_groups=None):
        """
        Disk-based transform of one worksheet part.
        Strategy: the file is processed as head (first chunk, holds <dimension>),
        middle (streamed through untouched), and tail (last region, holds the
        final rows + </sheetData> + autoFilter). set_cells requires a full parse,
        so it is only allowed on sheets small enough to hold in memory (< 32MB
        decompressed) — Dashboard-class sheets. Appends work on any size.
        """
        size = os.path.getsize(src)
        changed = 0

        if set_cells or paste_groups:
            if size > 32 * 1024 * 1024:
                raise ValueError("set_cells only supported on sheets < 32MB "
                                 "decompressed; use append_rows for data tabs")
            with open(src, "r", encoding="utf-8") as f:
                xml = f.read()
            # ONE merged extract/rebuild/join for whatever combination of
            # paste_columns + set_cells targets this sheet, instead of one
            # full pass per op — see _apply_row_ops. On a real month's file
            # this is the difference between two full in-memory copies of
            # 'New Sales report' and four.
            if paste_groups or set_cells:
                xml, n = self._apply_row_ops(
                    xml, paste_groups or [], [set_cells] if set_cells else [])
                changed += n
            if append_groups:
                xml, n = self._apply_append_inmem(xml, append_groups)
                changed += n
            with open(dst, "w", encoding="utf-8") as f:
                f.write(xml)
            return changed

        # append-only path: bounded memory regardless of sheet size
        tail_len = min(size, 2 * 1024 * 1024)
        with open(src, "rb") as f:
            f.seek(size - tail_len)
            tail = f.read().decode("utf-8")

        close_idx = tail.rfind("</sheetData>")
        if close_idx == -1:
            m = re.search(r"<sheetData\s*/>", tail)
            if not m:
                raise ValueError("</sheetData> not found in tail window; "
                                 "increase window")
            tail = tail[:m.start()] + "<sheetData></sheetData>" + tail[m.end():]
            close_idx = tail.rfind("</sheetData>")

        last_row = 0
        for m in re.finditer(r'<row r="(\d+)"', tail[:close_idx]):
            last_row = max(last_row, int(m.group(1)))

        new_rows_xml, appended, ncols_new = [], 0, 0
        r = last_row
        for rows, styles in append_groups:
            for values in rows:
                r += 1
                new_rows_xml.append(row_xml(r, values, styles))
                ncols_new = max(ncols_new, len(values))
                appended += 1
                changed += len(values)
        insertion = "".join(new_rows_xml)
        tail = tail[:close_idx] + insertion + tail[close_idx:]

        # extend autoFilter / dimension end refs that sit in the tail
        def extend_ref(mm):
            start, end = mm.group(1), mm.group(2)
            ec, _ = split_ref(end)
            ec_i = max(col_index(ec), ncols_new)
            return f'ref="{start}:{col_letter(ec_i)}{r}"'
        tail = re.sub(r'ref="([A-Z]+\d+):([A-Z]+\d+)"', extend_ref,
                      tail[close_idx + len(insertion):]) \
            .join([tail[:close_idx + len(insertion)], ""]) if False else tail
        tail = re.sub(r'(<autoFilter[^>]*ref=")([A-Z]+\d+):([A-Z]+\d+)(")',
                      lambda m: f'{m.group(1)}{m.group(2)}:'
                                f'{m.group(3)[:re.match(r"[A-Z]+", m.group(3)).end()]}{r}{m.group(4)}',
                      tail)

        head_len = min(size - tail_len, 64 * 1024) if size > tail_len else 0
        with open(src, "rb") as fin, open(dst, "wb") as fout:
            if head_len:
                head = fin.read(head_len).decode("utf-8", errors="surrogateescape")
                head = re.sub(
                    r'(<dimension ref="[A-Z]+\d+):([A-Z]+)(\d+)(")',
                    lambda m: f'{m.group(1)}:'
                              f'{col_letter(max(col_index(m.group(2)), ncols_new))}{r}{m.group(4)}',
                    head, count=1)
                fout.write(head.encode("utf-8", errors="surrogateescape"))
                remaining = size - head_len - tail_len
                while remaining > 0:
                    buf = fin.read(min(CHUNK, remaining))
                    if not buf:
                        break
                    fout.write(buf)
                    remaining -= len(buf)
            else:
                # whole file fit in the tail window
                pass
            fout.write(tail.encode("utf-8"))
        return changed

    def _apply_set_cells(self, xml: str, cells: dict) -> tuple[str, int]:
        return self._apply_row_ops(xml, [], [cells])

    @staticmethod
    def _rebuild_row(row_frag: str, row_num: int, colvals: dict) -> str:
        header = re.match(r'<row[^>]*>', row_frag).group(0)
        header = re.sub(r'\sspans="[^"]*"', "", header)  # stale spans are worse than none
        if header.endswith("/>"):
            existing = []
            header = header[:-2] + ">"
        else:
            inner = row_frag[len(re.match(r'<row[^>]*>', row_frag).group(0)):-len("</row>")]
            existing = re.findall(r'<c\b[^>]*(?:/>|>.*?</c>)', inner, re.S)
        cells: dict[int, str] = {}
        for cx in existing:
            ref = re.search(r'r="([A-Z]+)\d+"', cx)
            if ref:
                cells[col_index(ref.group(1))] = cx
        for col, val in colvals.items():
            x = cell_xml(f"{col}{row_num}", val)
            ci = col_index(col)
            if x:
                cells[ci] = x
            else:
                cells.pop(ci, None)
        body = "".join(cells[k] for k in sorted(cells))
        return header + body + "</row>"

    def _apply_row_ops(self, xml: str, paste_groups: list,
                       cell_groups: list) -> tuple[str, int]:
        """Single streaming pass — never materializes a dict of every row's
        content — for however many paste_columns/set_cells ops target ONE
        sheet part.

        An earlier version built `{row_num: full_row_fragment}` via a dict
        comprehension over the WHOLE document before touching anything. That
        forces Python's lazy `finditer` iterator to fully materialize —
        every row's content copied a second time into new string objects,
        on top of the source string already held in memory. On a real
        month's file that's exactly AR_07.31 (paste A:X + set_cells Y:AH,
        ~16k rows) and 'New Sales report' (paste A:Y + set_cells Z:AH,
        ~13.5k rows) — the two sheets that OOM'd a 512MB Render instance.

        This version iterates `finditer` lazily, one row at a time (same
        technique the disk-streamed append path already uses successfully
        on a sheet with 464,908+ rows), deciding immediately whether to
        rebuild or pass a row through, and discarding the match before
        moving to the next. Peak memory is roughly source + result (~2x
        the sheet size), not existing-dict + result-dict + source + join
        (~3-4x). New rows past the source's extent, and clear_beyond on
        rows below the pasted block, are handled inline via a sorted
        pointer into the touched-row list — relies on OOXML rows always
        being stored in ascending r= order (the append path already
        assumes this same invariant).
        """
        xml = re.sub(r"<sheetData\s*/>", "<sheetData></sheetData>", xml, count=1)
        open_m = re.search(r"<sheetData>", xml)
        close_i = xml.rfind("</sheetData>")
        if not open_m or close_i == -1:
            raise ValueError("sheetData not found")
        head, body, tail = xml[:open_m.end()], xml[open_m.end():close_i], xml[close_i:]

        row_colvals: dict[int, dict[str, object]] = {}
        clear_specs = []      # (block_cols, last_new_row) per paste op
        for anchor, rows, clear_beyond in paste_groups:
            a_col, a_row = split_ref(anchor)
            start_col = col_index(a_col)
            width = max((len(r) for r in rows), default=0)
            for i, values in enumerate(rows):
                rn = a_row + i
                colvals = row_colvals.setdefault(rn, {})
                for j in range(width):
                    colvals[col_letter(start_col + j)] = (
                        values[j] if j < len(values) else None)
            if clear_beyond and width:
                clear_specs.append((set(range(start_col, start_col + width)),
                                    a_row + len(rows) - 1))

        write_refs = []
        for cells in cell_groups:
            for ref, val in cells.items():
                col, row = split_ref(ref)
                row_colvals.setdefault(row, {})[col] = val
                write_refs.append((ref, val))

        changed = sum(len(cv) for cv in row_colvals.values())

        pending = sorted(row_colvals)   # touched rows, ascending; may or may
        pi = 0                          # not exist in the source
        out_pieces = []

        def flush_new_before(limit):
            nonlocal pi
            while pi < len(pending) and pending[pi] < limit:
                rn = pending[pi]
                out_pieces.append(self._rebuild_row(
                    f'<row r="{rn}"></row>', rn, row_colvals[rn]))
                pi += 1

        row_re = re.compile(r'<row r="(\d+)"(?:\s[^>]*)?(?:/>|>.*?</row>)', re.S)
        for m in row_re.finditer(body):
            rn = int(m.group(1))
            flush_new_before(rn)   # any pending rows strictly before rn are
                                   # genuinely missing from the source
            colvals = row_colvals.get(rn)
            if colvals is not None:
                out_pieces.append(self._rebuild_row(m.group(0), rn, colvals))
                if pi < len(pending) and pending[pi] == rn:
                    pi += 1        # this row existed after all — consumed
                continue
            cleared = False
            for block_cols, last_new_row in clear_specs:
                if rn > last_new_row:
                    out_pieces.append(self._rebuild_row(
                        m.group(0), rn, {c: None for c in map(col_letter, block_cols)}))
                    changed += len(block_cols)
                    cleared = True
                    break
            if not cleared:
                out_pieces.append(m.group(0))
        flush_new_before(float("inf"))   # remaining rows past the last existing one

        new_body = "".join(out_pieces)
        xml = head + new_body + tail

        # every set_cells ref we wrote a value into must actually be present
        # (None/"" deletes the cell, so those refs are exempt) — paste values
        # are not checked here, matching the prior per-op behavior
        missing = [ref for ref, val in write_refs
                   if val not in (None, "") and f'r="{ref}"' not in xml]
        if missing:
            raise ValueError(f"set_cells failed to write {missing} — the "
                             "cells are not present in the transformed sheet")
        return xml, changed

    def _apply_paste(self, xml: str, anchor: str, rows: list,
                     clear_beyond: bool) -> tuple[str, int]:
        return self._apply_row_ops(xml, [(anchor, rows, clear_beyond)], [])

    def _apply_append_inmem(self, xml: str, append_groups) -> tuple[str, int]:
        last_row = 0
        for m in re.finditer(r'<row r="(\d+)"', xml):
            last_row = max(last_row, int(m.group(1)))
        r = last_row
        rows_xml, changed = [], 0
        for rows, styles in append_groups:
            for values in rows:
                r += 1
                rows_xml.append(row_xml(r, values, styles))
                changed += len(values)
        return (xml.replace("</sheetData>", "".join(rows_xml) + "</sheetData>", 1),
                changed)
