"""Append-path memory lock-in.

Job 2ba2c030cd80 (2026-08-19) OOM'd Render at 512MB inside phase 2: the
append path built one giant `insertion` string, concatenated it into the
tail, then re.sub'd the merged copy — ~160MB of stacked transients for the
real 13,500-row 'Sales report Raw' append. The path now streams rows out
in 512-row batches. This test appends a real-shaped batch (13,500 rows,
25 data cols + 2 formula cells, inline strings) to a ~40MB synthetic sheet
and asserts both correctness and a bounded Python-level allocation peak.
"""
import os, re, sys, tempfile, tracemalloc, zipfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:280]}"))
    if not cond:
        fails.append(label)


TD = tempfile.mkdtemp(prefix="appmem_")

# ---- synthetic 'Sales report Raw'-shaped part: ~300k rows, ~40MB
src_part = os.path.join(TD, "raw_src.xml")
N_EXIST = 300_000
with open(src_part, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:AA{N_EXIST}"/><sheetData>')
    for r in range(1, N_EXIST + 1):
        f.write(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>Client {r % 997}</t></is></c>'
                f'<c r="L{r}"><v>{r * 1.5}</v></c>'
                f'<c r="Y{r}" t="inlineStr"><is><t>Company {r % 89} LLC</t></is></c></row>')
    f.write(f'</sheetData><autoFilter ref="A3:AA{N_EXIST}"/></worksheet>')
size_mb = os.path.getsize(src_part) / 1048576
print(f"synthetic source part: {size_mb:.1f} MB, {N_EXIST} rows")

# ---- real-shaped append: 13,500 rows x 25 data cols + 2 formula cells
rows = []
for i in range(13_500):
    row = [f"Client {i}", "Wholesale", "Romane", "2024-01-15", "Store",
           "2026-07-10", "CustInvc", f"INV{i}", f"SKU-{i % 400}",
           f"Eau de Parfum {i % 50} 100ml spray with box", 3, 145.5 + i % 7,
           "EMP101 Mark Saviski", "TX", f"Partner {i % 20}", "", "10", "",
           "", "", "", "", "Jul 2026", 19.71, f"Company {i % 89} LLC"]
    row += ["=V{r}-L{r}", "=Z{r}*0.1"]      # Z, AA formula templates
    rows.append(row)

s = XlsxSurgeon.__new__(XlsxSurgeon)       # only _transform_sheet is exercised
s.workdir = TD
dst_part = os.path.join(TD, "raw_out.xml")

tracemalloc.start()
changed = s._transform_sheet(src_part, dst_part, {}, [(rows, None)])
cur, peak = tracemalloc.get_traced_memory()
tracemalloc.stop()
peak_mb = peak / 1048576
print(f"append done: {changed} cells changed, python peak {peak_mb:.1f} MB")

check("python allocation peak bounded (< 80MB)", peak_mb < 80, f"{peak_mb:.1f} MB")
check("changed count = rows x width", changed == 13_500 * 27, changed)

with open(dst_part, encoding="utf-8") as f:
    head = f.read(4096)
    f.seek(os.path.getsize(dst_part) - 24 * 1024 * 1024)
    tail = f.read()
last = N_EXIST + 13_500
check("dimension extended to final row+width",
      f'<dimension ref="A1:AA{last}"/>' in head, head[:200])
check("autoFilter extended", f'<autoFilter ref="A3:AA{last}"/>' in tail)
check("first appended row present with formula",
      f'<row r="{N_EXIST + 1}">' in tail
      and f'<c r="Z{N_EXIST + 1}"><f>V{N_EXIST + 1}-L{N_EXIST + 1}</f></c>' in tail)
check("last appended row present",
      f'<c r="AA{last}"><f>Z{last}*0.1</f></c>' in tail)
check("rows in order, no duplicates",
      tail.count(f'<row r="{last}"') == 1
      and tail.find(f'<row r="{last - 1}"') < tail.find(f'<row r="{last}"'))
check("</sheetData> after last row, exactly once",
      tail.count("</sheetData>") == 1
      and tail.find(f'<row r="{last}"') < tail.find("</sheetData>"))
check("last kept source row survives", f'<row r="{N_EXIST}">' in tail)

for p in (src_part, dst_part):
    try:
        os.remove(p)
    except OSError:
        pass

print()
if fails:
    print(f"{len(fails)} FAILURES: {fails}")
    sys.exit(1)
print("ALL PASS")
