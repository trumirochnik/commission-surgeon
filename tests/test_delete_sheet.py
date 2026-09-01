"""delete_sheet op (Mike/Preet review 2026-08-25: his manual flow recycles
two AR tabs; ours left the two-months-back tab in place — he reviewed the
stale tab, confused. The op removes it AFTER the retargets, with a guard
that refuses while anything still references the sheet)."""
import os
import re
import sys
import tempfile
import zipfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:260]}"))
    if not cond:
        fails.append(label)


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WS_CT = ("application/vnd.openxmlformats-officedocument"
         ".spreadsheetml.worksheet+xml")


def sheet_xml(formula=None, text=None):
    cell = ""
    if formula:
        cell = f'<c r="A1"><f>{formula}</f><v>0</v></c>'
    elif text:
        cell = f'<c r="A1" t="inlineStr"><is><t>{text}</t></is></c>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{NS_MAIN}"><sheetData>'
            f'<row r="1">{cell}</row></sheetData></worksheet>')


def build(path, ref_formula=None, defined_name=None, shared_text=None):
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}">'
          '<bookViews><workbookView activeTab="2"/></bookViews>'
          '<sheets>'
          '<sheet name="Keep" sheetId="1" r:id="rId1"/>'
          '<sheet name="Mid" sheetId="2" r:id="rId2"/>'
          '<sheet name="AR_05.31" sheetId="3" r:id="rId3"/>'
          '</sheets>'
          + (f'<definedNames><definedName name="X">{defined_name}'
             '</definedName></definedNames>' if defined_name else "")
          + '<calcPr calcId="1"/></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<Relationships xmlns="{NS_REL}">'
               f'<Relationship Id="rId1" Type="{NS_R}/worksheet" Target="worksheets/sheet1.xml"/>'
               f'<Relationship Id="rId2" Type="{NS_R}/worksheet" Target="worksheets/sheet2.xml"/>'
               f'<Relationship Id="rId3" Type="{NS_R}/worksheet" Target="worksheets/sheet3.xml"/>'
               '</Relationships>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<Types xmlns="{NS_CT}">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{WS_CT}"/>'
          f'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="{WS_CT}"/>'
          f'<Override PartName="/xl/worksheets/sheet3.xml" ContentType="{WS_CT}"/>'
          '</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<Relationships xmlns="{NS_REL}">'
                 f'<Relationship Id="rId1" Type="{NS_R}/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml(formula=ref_formula))
        z.writestr("xl/worksheets/sheet2.xml", sheet_xml(text=shared_text))
        z.writestr("xl/worksheets/sheet3.xml", sheet_xml(text="stale May data"))


tmp = tempfile.mkdtemp(prefix="delsheet_")
SRC = os.path.join(tmp, "src.xlsx")
DST = os.path.join(tmp, "out.xlsx")

# ── 1. clean delete ──
build(SRC)
s = XlsxSurgeon(SRC, workdir=tmp)
s.delete_sheet("AR_05.31")
res = s.apply(DST)
with zipfile.ZipFile(DST) as z:
    names = z.namelist()
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    ct = z.read("[Content_Types].xml").decode()
check("D1: part gone from the zip", "xl/worksheets/sheet3.xml" not in names)
check("D2: workbook entry gone", 'name="AR_05.31"' not in wb
      and len(re.findall(r"<sheet\b", wb)) == 2, wb)
check("D3: relationship gone", 'Id="rId3"' not in rels)
check("D4: content-type override gone", "sheet3.xml" not in ct)
check("D5: out-of-range activeTab clamped", 'activeTab="0"' in wb, wb)
check("D6: results entry present",
      any(r.get("op") == "delete_sheet" and r.get("sheet") == "AR_05.31"
          and r.get("cellsChanged") == 1 for r in res), res)
s2 = XlsxSurgeon(DST, workdir=tmp)
check("D7: output remaps to exactly the two kept sheets",
      sorted(s2.sheet_names()) == ["Keep", "Mid"], s2.sheet_names())

# ── 2. guard: live formula reference refuses ──
build(SRC, ref_formula="SUM('AR_05.31'!A1:A9)")
s = XlsxSurgeon(SRC, workdir=tmp)
s.delete_sheet("AR_05.31")
try:
    s.apply(DST)
    check("G1: referenced sheet refuses to delete", False, "no raise")
except ValueError as e:
    check("G1: referenced sheet refuses to delete",
          "still referenced" in str(e) and "sheet1" in str(e), e)

# ── 3. guard: defined name refuses ──
build(SRC, defined_name="'AR_05.31'!$A$1")
s = XlsxSurgeon(SRC, workdir=tmp)
s.delete_sheet("AR_05.31")
try:
    s.apply(DST)
    check("G2: defined-name reference refuses", False, "no raise")
except ValueError as e:
    check("G2: defined-name reference refuses", "defined name" in str(e), e)

# ── 4. plain cell TEXT naming the sheet does NOT block (labels are fine) ──
build(SRC, shared_text="see tab AR_05.31 for May")
s = XlsxSurgeon(SRC, workdir=tmp)
s.delete_sheet("AR_05.31")
res = s.apply(DST)
check("G3: cell text mentioning the name doesn't block",
      any(r.get("op") == "delete_sheet" for r in res), res)

# ── 5. delete + other op on the same sheet refuses ──
build(SRC)
s = XlsxSurgeon(SRC, workdir=tmp)
s.set_cells("AR_05.31", {"A1": "x"})
s.delete_sheet("AR_05.31")
try:
    s.apply(DST)
    check("G4: delete+edit same sheet refuses", False, "no raise")
except ValueError as e:
    check("G4: delete+edit same sheet refuses", "also targeted" in str(e), e)

# ── 6. unknown sheet refuses at queue time ──
build(SRC)
s = XlsxSurgeon(SRC, workdir=tmp)
try:
    s.delete_sheet("Nope")
    check("G5: unknown sheet refuses", False, "no raise")
except ValueError as e:
    check("G5: unknown sheet refuses", "not found" in str(e), e)

print()
print("ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
