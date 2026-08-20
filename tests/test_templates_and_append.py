"""Tests for the newly-filled FORMULA_TEMPLATES: prior_ar_tab substitution,
and the append_rows {r}-at-write-time fix for the 'raw' tab templates."""
import os, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from netsuite_extract import formula_cells, FORMULA_TEMPLATES
from surgeon import XlsxSurgeon
import commission_job

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {detail}"))
    if not cond:
        fails.append(label)


# ── T1: AR templates need prior_ar_tab, raise without it ──
try:
    formula_cells("ar", 7, 2)
    check("T1: missing prior_ar_tab raises", False)
except ValueError as e:
    check("T1: missing prior_ar_tab raises", "prior_ar_tab" in str(e))

# ── T2: with prior_ar_tab, AD/AG resolve to the real tab, not AR_05.31 ──
cells = formula_cells("ar", 7, 1, prior_ar_tab="AR_06.30")
check("T2a: AD7 references the real prior tab",
      "'AR_06.30'!AC:AC" in cells["AD7"] and "AR_05.31" not in cells["AD7"],
      cells["AD7"])
check("T2b: AG7 references the real prior tab",
      "'AR_06.30'!AC:AC" in cells["AG7"] and "AR_05.31" not in cells["AG7"],
      cells["AG7"])
check("T2c: no leftover {r} or {prior_ar_tab} placeholders anywhere",
      all("{r}" not in v and "{prior_ar_tab}" not in v for v in cells.values()),
      cells)
check("T2d: Z7 unaffected (no prior_ar_tab placeholder)", cells["Z7"] == "=V7-L7")

# ── T3: sales/raw templates need no prior_ar_tab (no placeholder inside) ──
sales_cells = formula_cells("sales", 7, 1)
check("T3: sales templates resolve without prior_ar_tab",
      all("{" not in v for v in sales_cells.values()), sales_cells)

# ── T4: build_ops end-to-end wires prior_ar_tab through from spec ──
data = {
    "arRows": [["c1", None, None, None, None, 44848, "Invoice", "T1", "I1",
               "d1", 1, 100.0] + [None] * 12],
    "salesRows": [["c1"] + [None] * 24],
}
spec = {
    "ar": {"target": "AR_07.31", "anchor": "A7", "formulaCols": "Y:AH"},
    "sales": {"target": "New Sales report", "anchor": "A7", "formulaCols": "Z:AH"},
    "raw": {"target": "Sales report Raw", "anchor": "A4", "formulaCols": "Z:AA"},
}
ops, report = commission_job.build_ops(data, spec, prior_ar_tab="AR_06.30")
ar_setcells_op = next(o for o in ops if o["op"] == "set_cells" and o["sheet"] == "AR_07.31")
check("T4a: build_ops AD7 uses the real prior tab",
      "AR_06.30" in ar_setcells_op["cells"]["AD7"], ar_setcells_op["cells"]["AD7"])
raw_append_op = next(o for o in ops if o["op"] == "append_rows")
check("T4b: raw append row STILL has literal {r} (not yet substituted)",
      "{r}" in raw_append_op["rows"][0][commission_job.col_to_index("Z") - 1],
      raw_append_op["rows"][0])

# ── T5: append_rows through the real surgeon substitutes {r} correctly ──
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
      '<sheets><sheet name="Sales report Raw" sheetId="1" r:id="rId1"/></sheets>'
      '<calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')
RAW = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
       '<dimension ref="A1:AA3"/><sheetData>'
       '<row r="3"><c r="A3" t="inlineStr"><is><t>Client</t></is></c></row>'
       '</sheetData></worksheet>')

WORK = tempfile.mkdtemp(prefix="tpl_test_")
src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", RAW)

tpl = FORMULA_TEMPLATES["raw"]
width = max(commission_job.col_to_index(c) for c in tpl)
rows = []
for i in range(3):
    row = [f"client{i}"] + [None] * (width - 1)
    for col, t in tpl.items():
        row[commission_job.col_to_index(col) - 1] = t
    rows.append(row)

s = XlsxSurgeon(src, workdir=WORK)
s.append_rows("Sales report Raw", rows)
out = os.path.join(WORK, "out.xlsx")
s.apply(out)
with zipfile.ZipFile(out) as z:
    raw_out = z.read("xl/worksheets/sheet1.xml").decode()

check("T5a: row 4 gets real row number in TEXT() formula",
      '<f>TEXT(T4,&quot;MMM YY&quot;)</f>' in raw_out, raw_out)
check("T5b: row 5 gets its OWN row number, not row 4's",
      '<f>TEXT(T5,&quot;MMM YY&quot;)</f>' in raw_out, raw_out)
check("T5c: row 6 (third appended row) also correct",
      '<f>TEXT(T6,&quot;MMM YY&quot;)</f>' in raw_out, raw_out)
check("T5d: no stray {r} literal survived into the file", "{r}" not in raw_out, raw_out)
check("T5e: AA column also correctly numbered per row",
      '<f>+CONCATENATE(H5,&quot; - &quot;,I5)</f>' in raw_out, raw_out)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
