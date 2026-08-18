"""Streaming duplicate rebuild at real scale: a ~34MB AR-like source sheet
(matching the measured real AR_06.30), full July-shaped op load (16,362-row
paste + 10 formula columns per row), verifying content, drop-beyond
semantics, and the memory win vs the in-memory path."""
import os, re, sys, time, tracemalloc, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {detail}"))
    if not cond:
        fails.append(label)


WORK = tempfile.mkdtemp(prefix="streamdup_")
N_SRC = 17000    # June had MORE rows than July's 16,362 -> tests drop-beyond
N_NEW = 16362

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

# ~2KB per source row so the whole part lands near the real 33.8MB
FAT = "June carryover content that pads this row to something like the real sheet " * 3


def src_row(r):
    cells = "".join(
        f'<c r="{chr(65 + c)}{r}" t="inlineStr"><is><t>{FAT[:80]}c{c}</t></is></c>'
        for c in range(20))
    cells += f'<c r="Y{r}"><f>IF(V{r}&gt;0,1,0)</f></c><c r="Z{r}"><f>V{r}-L{r}</f></c>'
    return f'<row r="{r}">{cells}</row>'


print(f"building a {N_SRC}-row source sheet...")
body_parts = ['<row r="6"><c r="A6" t="inlineStr"><is><t>HeaderClient</t></is></c></row>']
body_parts += [src_row(r) for r in range(7, 7 + N_SRC)]
SRC_SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             f'<dimension ref="A1:AH{6 + N_SRC}"/><sheetData>{"".join(body_parts)}</sheetData>'
             '<autoFilter ref="A6:X17006"/>'
             '<hyperlinks><hyperlink ref="A7" r:id="rId9"/></hyperlinks>'
             '<pageSetup orientation="landscape" r:id="rId1"/>'
             '</worksheet>')
print(f"source sheet: {len(SRC_SHEET) / 1_000_000:.1f} MB of text")
del body_parts

src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SRC_SHEET)
del SRC_SHEET

# July-shaped ops: 16,362 rows x 24 data cols + 10 formula cols per row
paste_rows = [[f"87{i:05d} July Store LLC", "Wholesale", "Romane", 44848,
               "Western", 44878, "Invoice", f"73{i:05d}", f"9{i:04d}",
               "Product description here", 12.0, 345.67, "EMP204 Kevin Hanks",
               "TX", f"Partner{i % 40}", "Primary Rep", None, "Invoice : Open",
               44968, None, "Primary", 345.67, None, "July Store LLC"]
              for i in range(N_NEW)]
cells = {}
for i in range(N_NEW):
    r = 7 + i
    for col in ("Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"):
        cells[f"{col}{r}"] = f'=IF(V{r}>0,VLOOKUP(I{r},\'Commission Rate by SKUs\'!B:E,4,0)," ")'

tracemalloc.start()
t0 = time.time()
s = XlsxSurgeon(src, workdir=WORK)
s.duplicate_sheet("AR_06.30", "AR_07.31")
s.paste_columns("AR_07.31", "A7", paste_rows, True)
s.set_cells("AR_07.31", cells)
out = os.path.join(WORK, "out.xlsx")
results = s.apply(out)
elapsed = time.time() - t0
_, peak = tracemalloc.get_traced_memory()
tracemalloc.stop()
peak_mb = peak / 1_000_000
print(f"streaming dup apply(): {elapsed:.1f}s, tracemalloc peak {peak_mb:.1f} MB")
print("results:", results)

dup_res = [r for r in results if r["target"] == "AR_07.31"][0]
check("S1: streaming path was used", dup_res.get("streamed") is True, dup_res)
# Python-level allocations must stay far below the source sheet size —
# the whole point is never materializing the 34MB part (note: the ops
# themselves — paste rows + formula dicts — dominate what remains)
check("S2: tracemalloc peak below 250MB (was 500+ in-memory equivalent)",
      peak_mb < 250, f"{peak_mb:.1f} MB")

with zipfile.ZipFile(out) as z:
    wb_out = z.read("xl/workbook.xml").decode()
    m = re.search(r'name="AR_07.31"[^>]*r:id="(rId\d+)"', wb_out)
    rels_out = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels_out).group(1)
    dup = z.read("xl/" + tgt).decode()

check("S3: header row 6 kept verbatim", "HeaderClient" in dup)
check("S4: first data row is July data",
      '87' + '00000' + ' July Store LLC' in dup and f'<row r="7">' in dup)
check("S5: last July row present", f'<row r="{6 + N_NEW}">' in dup)
check("S6: June rows beyond July extent DROPPED",
      f'<row r="{7 + N_NEW}">' not in dup and f'<row r="{6 + N_SRC}">' not in dup)
check("S7: formula columns present on first and last rows",
      dup.count("VLOOKUP") == N_NEW * 10
      and f"IF(V{6 + N_NEW}&gt;0" in dup, dup[-500:])
check("S8: no r:id anywhere in the copy", "r:id" not in dup)
check("S9: pageSetup/hyperlinks stripped",
      "<pageSetup" not in dup and "<hyperlink" not in dup)
check("S10: dimension updated to new extent",
      f'<dimension ref="A1:AH{6 + N_NEW}"/>' in dup, dup[:300])
check("S11: suffix (autoFilter) survived", "<autoFilter" in dup)
check("S12: June carryover text gone from data region", FAT[:40] not in dup)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
