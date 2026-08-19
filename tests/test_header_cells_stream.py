"""Header cells through the streaming rebuild (the AC6/AD4 month roll).

The 0819-2034 audit found two month-linked serials the roll missed, both
in header rows ABOVE the paste anchor: 'New Sales report'!AC6 (period end;
gates the AB deposit column that feeds Dashboard E) and AR!AD4 (month
start; gates the prior-month XLOOKUP branch in 16,358 AD formulas).
Streaming used to refuse cells outside the pasted spans; header cells are
now applied to the captured prefix, style-preserved and diff-aware.
"""
import os, re, sys, tempfile, zipfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:280]}"))
    if not cond:
        fails.append(label)


TD = tempfile.mkdtemp(prefix="hdrcell_")
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
      '<sheets><sheet name="AR_06.30" sheetId="1" r:id="rId1"/></sheets>'
      '<calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')
# header rows 1/4/6 (row 4 carries the month-start serial, date-styled),
# data rows 7-9 that the paste regenerates; NO row 2/3/5 (sparse header)
SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
         '<dimension ref="A1:AH9"/><sheetData>'
         '<row r="1"><c r="A1" t="inlineStr"><is><t>Aging</t></is></c></row>'
         '<row r="4"><c r="AD4" s="7"><v>46174</v></c></row>'
         '<row r="6"><c r="A6" t="inlineStr"><is><t>Client:Project</t></is></c></row>'
         '<row r="7"><c r="A7" t="inlineStr"><is><t>OldCo</t></is></c><c r="L7"><v>1</v></c></row>'
         '<row r="8"><c r="A8" t="inlineStr"><is><t>OldCo2</t></is></c></row>'
         '<row r="9"><c r="A9" t="inlineStr"><is><t>OldCo3</t></is></c></row>'
         '</sheetData></worksheet>')

src = os.path.join(TD, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)

# ── 1. dup path (the AR_07.31 case): paste + formulas + header serial
dst = os.path.join(TD, "out.xlsx")
s = XlsxSurgeon(src, workdir=TD)
s.duplicate_sheet("AR_06.30", "AR_07.31")
s.paste_columns("AR_07.31", "A7", [["NewCo", 10], ["NewCo2", 20]], clear_beyond=True)
s.set_cells("AR_07.31", {"Y7": "=L7*2", "Y8": "=L8*2",
                         "AD4": 46205,        # month start — prefix row
                         "B2": "inserted"})   # absent header row — insert in order
results = s.apply(dst)
dup = next(r for r in results if r["op"] == "duplicate_sheet")
check("dup path still STREAMS with header cells", dup.get("streamed") is True, dup)
with zipfile.ZipFile(dst) as z:
    parts = {n: z.read(n).decode() for n in z.namelist() if "worksheets" in n}
new = next(x for x in parts.values() if "NewCo" in x)
old = next(x for x in parts.values() if "OldCo</t>" in x and "NewCo" not in x)
check("AD4 rolled on the dup, style preserved",
      '<c r="AD4" s="7"><v>46205</v></c>' in new, new[:400])
check("source sheet AD4 untouched", '<c r="AD4" s="7"><v>46174</v></c>' in old)
check("absent header row inserted in row order",
      re.search(r'<row r="1">.*<row r="2"><c r="B2"[^>]*>.*<row r="4">', new, re.S), new[:500])
check("data region regenerated (old rows dropped)",
      "OldCo3" not in new and "NewCo2" in new and "<f>L7*2</f>" in new)
check("header label row 6 survives via prefix", "Client:Project" in new)

# ── 2. existing-sheet path (the New Sales AC6 case)
src2 = os.path.join(TD, "src2.xlsx")
with zipfile.ZipFile(src2, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB.replace("AR_06.30", "New Sales report"))
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml",
               SHEET.replace('<c r="AD4" s="7"><v>46174</v></c>',
                             '<c r="AC6" s="9"><v>46203</v></c>'
                             ).replace('<row r="4">', '<row r="6x">'
                             ).replace('6x', '6').replace(
                             '<row r="6"><c r="A6"', '<row r="5"><c r="A5"'))
dst2 = os.path.join(TD, "out2.xlsx")
s2 = XlsxSurgeon(src2, workdir=TD)
s2.paste_columns("New Sales report", "A7", [["JulyRow", 1]], clear_beyond=True)
s2.set_cells("New Sales report", {"Z7": "=B7*3", "AC6": 46234})
s2.apply(dst2)
with zipfile.ZipFile(dst2) as z:
    x2 = z.read("xl/worksheets/sheet1.xml").decode()
check("existing-sheet stream: AC6 rolled with style",
      '<c r="AC6" s="9"><v>46234</v></c>' in x2, x2[:400])
check("existing-sheet data + formula landed",
      "JulyRow" in x2 and "<f>B7*3</f>" in x2 and "OldCo3" not in x2)

print()
if fails:
    print(f"{len(fails)} FAILURES: {fails}")
    sys.exit(1)
print("ALL PASS")
