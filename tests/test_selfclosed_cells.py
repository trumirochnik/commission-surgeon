"""Self-closed cell/row parsing lock-in.

The 0820-1933 disaster: _CELL_RE's greedy paired branch let ([^>]*) swallow
the '/' of an empty cell (<c r="T7" s="4"/>), so its lazy body consumed the
ENTIRE NEXT CELL — T "read" U's shared-string index as a number and U
vanished. On the hand-built AR tabs (empty cells everywhere) that
misattributed the partner column on ~17,900 rows, and the v30 enrichment
then wrote the corruption back into the workbook. Same greedy-attr pattern
existed in the row regexes (self-closed spacer rows swallowed their
successor). These cases are the exact row-7 XML from the real June tab.
"""
import sys

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
import xlsx_read as xr

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:240]}"))
    if not cond:
        fails.append(label)


# the real June-tab row 7 shape: empty self-closed T followed by shared-string U
ROW = ('<row r="7"><c r="S7" s="4"><v>44811</v></c>'
       '<c r="T7" s="4"/><c r="U7" s="3" t="s"><v>52</v></c>'
       '<c r="V7" s="166"><v>-22.5</v></c>'
       '<c r="W7" s="3" t="s"><v>1316</v></c></row>')
cells = xr.parse_row_cells(ROW)
check("empty self-closed T stays empty",
      "T" in cells and cells["T"]["v"] is None and cells["T"]["t"] is None, cells.get("T"))
check("U keeps its shared-string index",
      cells.get("U", {}).get("t") == "s" and cells["U"]["v"] == "52", cells.get("U"))
check("V numeric intact", cells.get("V", {}).get("v") == "-22.5", cells.get("V"))
check("all five cells parsed", set("STUVW") <= set(cells), sorted(cells))

# self-closed cell as the LAST cell of the row
ROW2 = '<row r="9"><c r="A9"><v>1</v></c><c r="B9" s="2"/></row>'
c2 = xr.parse_row_cells(ROW2)
check("trailing self-closed cell parsed empty",
      c2.get("A", {}).get("v") == "1" and "B" in c2 and c2["B"]["v"] is None, c2)

# self-closed spacer ROW must not swallow its successor (stream_rows path)
import io, zipfile, tempfile, os
TD = tempfile.mkdtemp(prefix="scr_")
SHEET = ('<?xml version="1.0"?><worksheet><sheetData>'
         '<row r="4" ht="15" customHeight="1"/>'
         '<row r="5"><c r="A5" t="inlineStr"><is><t>alive</t></is></c></row>'
         '</sheetData></worksheet>')
p = os.path.join(TD, "t.zip")
with zipfile.ZipFile(p, "w") as z:
    z.writestr("xl/worksheets/sheet1.xml", SHEET)
with zipfile.ZipFile(p) as z:
    rows = xr.stream_rows(z, "xl/worksheets/sheet1.xml", 1, 10)
check("self-closed row parsed on its own", 4 in rows and rows[4] == {}, rows.get(4))
check("successor row survives with its cells",
      rows.get(5, {}).get("A", {}).get("is") == "alive", rows.get(5))

print()
if fails:
    print(f"{len(fails)} FAILURES: {fails}")
    sys.exit(1)
print("ALL PASS")
