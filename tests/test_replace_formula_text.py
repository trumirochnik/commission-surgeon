"""replace_formula_text (the Dashboard date roll) tests.

Mirrors the real July'26 Dashboard: 'AR_05.31'!$T:$T,"<46204" receipt
cutoffs stored as "&lt;46204", four trailing manual credit-memo constants
(+H6-1335.11 style), shared-formula followers, and an &-containing
criteria formula that must round-trip byte-identically.
"""
import os, re, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:280]}"))
    if not cond:
        fails.append(label)


TD = tempfile.mkdtemp(prefix="freplace_")
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheets><sheet name="Dashboard" sheetId="1" r:id="rId1"/></sheets>'
      '<calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')

# E6/E7: cutoff literal + trailing manual constants; E8: &-criteria formula
# that is NOT targeted; E9: shared master + self-closed follower; C33: a
# date-styled VALUE cell (s="7") that set_cells must restyle-preserve;
# B1: a label containing 46204 that must NOT be touched (it's not a formula)
DASH = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:X40"/><sheetData>'
        '<row r="1"><c r="B1" t="inlineStr"><is><t>cutoff was 46204</t></is></c>'
        '<c r="C1"><v>46204</v></c></row>'
        '<row r="6"><c r="E6"><f>SUMIFS(\'AR_05.31\'!$L:$L,\'AR_05.31\'!$T:$T,"&lt;46204")+H6-1335.11</f><v>1</v></c></row>'
        '<row r="7"><c r="E7"><f>SUMIFS(\'AR_05.31\'!$T:$T,"&lt;46204")+H7+1845.04</f><v>2</v></c></row>'
        '<row r="8"><c r="E8"><f>SUMIFS(X:X,AA:AA,"&lt;&gt;"&amp;" ")</f><v>3</v></c></row>'
        '<row r="9"><c r="E9"><f t="shared" ref="E9:E10" si="0">T9*2</f><v>4</v></c></row>'
        '<row r="10"><c r="E10"><f t="shared" si="0"/><v>5</v></c></row>'
        '<row r="33"><c r="C33" s="7"><v>46173</v></c></row>'
        '</sheetData></worksheet>')


def make_src(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WB)
        z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
        z.writestr("xl/worksheets/sheet1.xml", DASH)


def sheet_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode()


# ── 1. the real date-roll combo: literal + regex strip + set_cells + retarget
src = os.path.join(TD, "src.xlsx")
dst = os.path.join(TD, "out.xlsx")
make_src(src)
s = XlsxSurgeon(src, workdir=TD)
s.retarget_refs("Dashboard", [{"from": "AR_05.31", "to": "AR_06.30"}])
s.replace_formula_text("Dashboard", [
    {"from": '"<46204"', "to": '"<46235"'},
    {"from": r"(\+H\d+)(?:[-+][0-9]+(?:\.[0-9]+)?)+$", "to": r"\1",
     "regex": True, "optional": True},
    {"from": "NEVER_PRESENT", "to": "x", "optional": True},
])
s.set_cells("Dashboard", {"C33": 46203})
results = s.apply(dst)
xml = sheet_xml(dst)

check("cutoff literal rolled in both formulas",
      xml.count('"&lt;46235"') == 2 and '"&lt;46204"' not in xml, xml[:600])
check("E6 manual constant stripped, +H6 kept",
      "+H6</f>" in xml and "-1335.11" not in xml, xml)
check("E7 plus-sign constant stripped too",
      "+H7</f>" in xml and "+1845.04" not in xml, xml)
check("&-criteria formula round-trips byte-identical",
      'SUMIFS(X:X,AA:AA,"&lt;&gt;"&amp;" ")' in xml, xml)
check("non-formula 46204s untouched (label + value)",
      "cutoff was 46204" in xml and "<v>46204</v>" in xml, xml)
check("shared follower <f/> untouched",
      '<f t="shared" si="0"/>' in xml, xml)
check("C33 value written with style preserved",
      re.search(r'<c r="C33" s="7"><v>46203</v></c>', xml), xml)
fr = next(r for r in results if r["op"] == "replace_formula_text")
check("perMapping counts: literal=2, strip=2, optional-miss=0",
      fr["perMapping"]['"<46204"'] == 2
      and fr["perMapping"][r"(\+H\d+)(?:[-+][0-9]+(?:\.[0-9]+)?)+$"] == 2
      and fr["perMapping"]["NEVER_PRESENT"] == 0, fr)
check("retarget still applied alongside",
      "'AR_06.30'!$L:$L" in xml and "AR_05.31" not in xml, xml[:600])

# ── 1b. self-closed follower BEFORE a target formula, with later </f>s:
# the 0821-1356 corruption shape (Compiled Data). The follower must not
# swallow the span, and the target must still be replaced cleanly.
make_src(src)
s = XlsxSurgeon(src, workdir=TD)
# fixture: row 9 has the master (paired), row 10 the self-closed follower,
# and E6/E7 targets come AFTER in document order? Build a dedicated sheet:
SHEET2 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<dimension ref="A1:X40"/><sheetData>'
          '<row r="5"><c r="G5" s="1"><f t="shared" ref="G5:G6" si="9">SUMIFS(X:X,"&lt;40074")</f><v>1</v></c>'
          '<c r="H5" s="1"><f>N5*$40074</f><v>2</v></c></row>'
          '<row r="6"><c r="G6" s="1"><f t="shared" si="9"/><v>3</v></c>'
          '<c r="H6" s="1"><v>44</v></c>'
          '<c r="I6" s="1"><f>SUMIFS(Y:Y,"&lt;40074")</f><v>5</v></c></row>'
          '</sheetData></worksheet>')
import zipfile as _zf2
with open(src, "rb") as f:
    pass
with _zf2.ZipFile(src, "w", _zf2.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET2)
s = XlsxSurgeon(src, workdir=TD)
s.replace_formula_text("Dashboard", [{"from": "40074", "to": "45999"}])
res_sc = s.apply(os.path.join(TD, "out_sc.xlsx"))
with _zf2.ZipFile(os.path.join(TD, "out_sc.xlsx")) as z:
    xsc = z.read("xl/worksheets/sheet1.xml").decode()
import xml.etree.ElementTree as _ET
try:
    _ET.fromstring(xsc)
    wellformed = True
except _ET.ParseError:
    wellformed = False
check("self-closed follower does not swallow the span (well-formed XML)",
      wellformed, xsc[:400])
check("all three targets replaced, follower untouched",
      xsc.count("45999") == 3 and '<f t="shared" si="9"/>' in xsc
      and "&lt;c" not in xsc and "40074" not in xsc, xsc)
fr_sc = next(r for r in res_sc if r["op"] == "replace_formula_text")
check("counts correct with followers present",
      fr_sc["perMapping"]["40074"] == 3, fr_sc)

# ── 2. non-optional mapping matching nothing fails the job
make_src(src)
s = XlsxSurgeon(src, workdir=TD)
s.replace_formula_text("Dashboard", [{"from": '"<99999"', "to": '"<11111"'}])
try:
    s.apply(os.path.join(TD, "out2.xlsx"))
    check("zero-match non-optional mapping refused", False, "no exception")
except ValueError as e:
    check("zero-match non-optional mapping refused", "matched nothing" in str(e), e)

# ── 3. validation errors
make_src(src)
s = XlsxSurgeon(src, workdir=TD)
try:
    s.replace_formula_text("Dashboard", [{"from": "", "to": "x"}])
    check("empty 'from' rejected", False, "no exception")
except ValueError as e:
    check("empty 'from' rejected", True)
try:
    s.replace_formula_text("Dashboard", [{"from": "([bad", "to": "x", "regex": True}])
    check("bad regex rejected at queue time", False, "no exception")
except re.error:
    check("bad regex rejected at queue time", True)
try:
    s.replace_formula_text("NoSuchSheet", [{"from": "a", "to": "b"}])
    check("unknown sheet rejected", False, "no exception")
except KeyError:
    check("unknown sheet rejected", True)

# ── 4. works on a sheet duplicated in the same job (in-memory fallback)
make_src(src)
s = XlsxSurgeon(src, workdir=TD)
s.duplicate_sheet("Dashboard", "Dash2")
s.replace_formula_text("Dash2", [{"from": '"<46204"', "to": '"<46235"'}])
results = s.apply(os.path.join(TD, "out3.xlsx"))
with zipfile.ZipFile(os.path.join(TD, "out3.xlsx")) as z:
    parts = {n: z.read(n).decode() for n in z.namelist()
             if n.startswith("xl/worksheets/")}
dup_part = next(x for n, x in parts.items() if '"&lt;46235"' in x)
orig_part = next(x for n, x in parts.items() if '"&lt;46204"' in x)
check("dup got the roll, original untouched",
      dup_part.count('"&lt;46235"') == 2 and orig_part.count('"&lt;46204"') == 2)

print()
if fails:
    print(f"{len(fails)} FAILURES: {fails}")
    sys.exit(1)
print("ALL PASS")
