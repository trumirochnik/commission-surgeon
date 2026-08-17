"""Stress-scale regression test for the O(n^2) set_cells bug.

The bug only shows up at realistic row counts: a 3-cell test (the original
unit tests) never exercises it. This builds a ~16,362-row sheet with
~1.5KB/row (roughly matching the real AR_06.30 sheet's data density) and
times a single set_cells call writing one formula cell per row — exactly
commission_job.py's AR Z-column pattern that stalled the real Render job.
"""
import os, sys, time, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

WORK = tempfile.mkdtemp(prefix="scale_test_")
N = 16362

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


def build_row(r):
    # ~24 data cols with realistic-length inline strings, mirroring real AR rows
    cells = [f'<c r="A{r}" t="inlineStr"><is><t>875{r:03d} Some Fixture Retail Store Name LLC</t></is></c>',
             f'<c r="I{r}" t="inlineStr"><is><t>Tru Western - Yellowstone Mens EDC 100mL Original</t></is></c>',
             f'<c r="L{r}"><v>{r * 1.23}</v></c>',
             f'<c r="V{r}"><v>{r * 4.56}</v></c>']
    return f'<row r="{r}">' + "".join(cells) + "</row>"


print(f"building a {N}-row fixture...")
body = "".join(build_row(r) for r in range(7, 7 + N))
AR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      f'<dimension ref="A1:X{6+N}"/><sheetData>{body}</sheetData></worksheet>')
print(f"fixture sheet size: {len(AR) / 1_000_000:.1f} MB")

src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", AR)

# exactly commission_job.py's pattern: one Z-column formula cell per AR row
cells = {f"Z{r}": f"=V{r}-L{r}" for r in range(7, 7 + N)}

s = XlsxSurgeon(src, workdir=WORK)
s.set_cells("AR_06.30", cells)

t0 = time.time()
out = os.path.join(WORK, "out.xlsx")
results = s.apply(out)
elapsed = time.time() - t0

print(f"apply() with {N} set_cells took {elapsed:.2f}s")
print("results:", results)

with zipfile.ZipFile(out) as z:
    ar_out = z.read("xl/worksheets/sheet1.xml").decode()

ok_time = elapsed < 20.0     # was 45+ MINUTES before the fix; generous CI margin
ok_correct = ('<c r="Z7"><f>V7-L7</f></c>' in ar_out
              and f'<c r="Z{6+N}"><f>V{6+N}-L{6+N}</f></c>' in ar_out
              and ar_out.count("<f>V") == N)

print("PASS" if ok_time else "FAIL", f"timing ({elapsed:.2f}s < 20s)")
print("PASS" if ok_correct else "FAIL", "correctness (first/last/count of formula cells)")
sys.exit(0 if (ok_time and ok_correct) else 1)
