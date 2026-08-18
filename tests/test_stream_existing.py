"""Streaming rebuild on an EXISTING sheet (the 'New Sales report' path):
content correct, prefix/suffix VERBATIM (r:id refs preserved — the sheet
keeps its _rels), rows beyond dropped, memory bounded at scale."""
import os, re, sys, time, tracemalloc, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {detail}"))
    if not cond:
        fails.append(label)


WORK = tempfile.mkdtemp(prefix="streamex_")
N_SRC = 14000
N_NEW = 13500

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
      '<sheets><sheet name="New Sales report" sheetId="1" r:id="rId1"/></sheets>'
      '<calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')
SHEET_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>'
              '</Relationships>')

FAT = "prior month sales row content padding out to realistic size " * 3


def src_row(r):
    cells = "".join(
        f'<c r="{chr(65 + c)}{r}" t="inlineStr"><is><t>{FAT[:70]}c{c}</t></is></c>'
        for c in range(20))
    return f'<row r="{r}">{cells}</row>'


print(f"building a {N_SRC}-row existing sheet...")
body = ('<row r="6"><c r="A6" t="inlineStr"><is><t>SalesHeader</t></is></c></row>'
        + "".join(src_row(r) for r in range(7, 7 + N_SRC)))
SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
         f'<dimension ref="A1:AH{6 + N_SRC}"/><sheetData>{body}</sheetData>'
         '<pageSetup orientation="landscape" r:id="rId1"/>'
         '</worksheet>')
print(f"sheet: {len(SHEET) / 1_000_000:.1f} MB of text")
del body

src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)
del SHEET

paste_rows = [[f"88{i:05d} July Sales Client", "Wholesale", "Romane", 44848,
               "Western", 44878, "Invoice", f"74{i:05d}", f"8{i:04d}",
               "Sales product desc", 6.0, 123.45, "EMP204 Kevin Hanks", "TX",
               f"P{i % 30}", "Primary Rep", None, "Invoice : Open", 44968,
               None, "Primary", 123.45, None, "July Sales Client", 1.25]
              for i in range(N_NEW)]
cells = {}
for i in range(N_NEW):
    r = 7 + i
    for col in ("Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"):
        cells[f"{col}{r}"] = f"=V{r}*0.05"

tracemalloc.start()
t0 = time.time()
s = XlsxSurgeon(src, workdir=WORK)
s.paste_columns("New Sales report", "A7", paste_rows, True)
s.set_cells("New Sales report", cells)
out = os.path.join(WORK, "out.xlsx")
results = s.apply(out)
elapsed = time.time() - t0
_, peak = tracemalloc.get_traced_memory()
tracemalloc.stop()
peak_mb = peak / 1_000_000
print(f"existing-sheet streaming apply(): {elapsed:.1f}s, tracemalloc peak {peak_mb:.1f} MB")
print("results:", results)

check("E1: memory bounded (<250MB python allocations)", peak_mb < 250, f"{peak_mb:.1f}")

with zipfile.ZipFile(out) as z:
    sheet = z.read("xl/worksheets/sheet1.xml").decode()
    rels_kept = "xl/worksheets/_rels/sheet1.xml.rels" in z.namelist()

check("E2: header row 6 kept", "SalesHeader" in sheet)
check("E3: July data present first and last",
      "8800000 July Sales Client" in sheet and f'<row r="{6 + N_NEW}">' in sheet)
check("E4: prior rows beyond new extent dropped",
      f'<row r="{7 + N_NEW}">' not in sheet)
check("E5: pageSetup r:id PRESERVED (existing sheet keeps its rels)",
      '<pageSetup orientation="landscape" r:id="rId1"/>' in sheet, sheet[-300:])
check("E6: sheet _rels part still present in output", rels_kept)
check("E7: formula cells written", sheet.count("*0.05") == N_NEW * 9)
check("E8: prior-month padding gone from data region", FAT[:40] not in sheet)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
