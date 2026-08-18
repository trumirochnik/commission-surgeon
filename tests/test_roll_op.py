"""copy_range_values (the SOP step 2 balance roll) tests."""
import os, re, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:280]}"))
    if not cond:
        fails.append(label)


TD = tempfile.mkdtemp(prefix="roll_")
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

# W6:W8 = formula cells with cached values (like the real ending balances),
# U6:U8 = last month's stale roll values; W7 blank tests None propagation
DASH = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:X30"/><sheetData>'
        '<row r="6"><c r="U6"><v>111.11</v></c>'
        '<c r="W6"><f>M6+N6+O6</f><v>16736.44</v></c></row>'
        '<row r="7"><c r="U7"><v>222.22</v></c></row>'
        '<row r="8"><c r="U8"><v>333.33</v></c>'
        '<c r="W8"><f>M8+N8+O8</f><v>28185.35</v></c></row>'
        '</sheetData></worksheet>')

src = os.path.join(TD, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", DASH)

out = os.path.join(TD, "out.xlsx")
s = XlsxSurgeon(src, workdir=TD)
s.copy_range_values("Dashboard", "W6:W8", "U6:U8")
s.set_cells("Dashboard", {"C2": "July'26"})
res = s.apply(out)
with zipfile.ZipFile(out) as z:
    dash = z.read("xl/worksheets/sheet1.xml").decode()

check("C1: U6 now holds W6's cached VALUE as a literal",
      '<c r="U6"><v>16736.44</v></c>' in dash, dash[:400])
check("C2: no formula copied into U", "<c r=\"U6\"><f>" not in dash)
check("C3: blank W7 clears U7 (roll writes the true state)",
      'r="U7"' not in dash, dash)
check("C4: U8 rolled too", '<c r="U8"><v>28185.35</v></c>' in dash)
check("C5: W column untouched (formulas + caches intact)",
      "<f>M6+N6+O6</f>" in dash and "<f>M8+N8+O8</f>" in dash)
check("C6: changes counted through the diff-aware path",
      sum(r["cellsChanged"] for r in res) >= 4, res)

# geometry mismatch refuses at queue time
s2 = XlsxSurgeon(src, workdir=TD)
try:
    s2.copy_range_values("Dashboard", "W6:W8", "U6:U9")
    check("C7: geometry mismatch raises", False)
except ValueError as e:
    check("C7: geometry mismatch raises", "geometry mismatch" in str(e), str(e))

# unbounded range refused
s3 = XlsxSurgeon(src, workdir=TD)
try:
    s3.copy_range_values("Dashboard", "W:W", "U:U")
    check("C8: unbounded range refused", False)
except ValueError as e:
    check("C8: unbounded range refused", "bounded" in str(e), str(e))

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
