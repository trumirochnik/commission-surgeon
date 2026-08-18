"""Two-phase surgery split: op classification + correctness vs single-phase."""
import os, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
os.environ.setdefault("NS_MCP_URL", "https://example.invalid/mcp")
import service
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {detail}"))
    if not cond:
        fails.append(label)


# ── classification logic (mirrors _run's inline split) ──
def classify(ops, ar_target):
    phase1, phase2 = [], []
    for op in ops:
        is_ar_or_dash = (
            (op["op"] == "duplicate_sheet" and op.get("name") == ar_target)
            or op.get("sheet") in (ar_target, "Dashboard"))
        (phase1 if is_ar_or_dash else phase2).append(op)
    return phase1, phase2


real_ops = [
    {"op": "duplicate_sheet", "source": "AR_06.30", "name": "AR_07.31"},
    {"op": "set_cells", "sheet": "Dashboard", "cells": {"C2": "July'26"}},
    {"op": "paste_columns", "sheet": "AR_07.31", "anchor": "A7", "rows": [[1]]},
    {"op": "set_cells", "sheet": "AR_07.31", "cells": {"Z7": "=V7-L7"}},
    {"op": "paste_columns", "sheet": "New Sales report", "anchor": "A7", "rows": [[1]]},
    {"op": "set_cells", "sheet": "New Sales report", "cells": {"Z7": "=Y7*2"}},
    {"op": "append_rows", "sheet": "Sales report Raw", "rows": [[1]]},
]
p1, p2 = classify(real_ops, "AR_07.31")
check("CL1: phase1 has exactly dup+dashboard+AR-paste+AR-setcells",
      [o["op"] for o in p1] == ["duplicate_sheet", "set_cells", "paste_columns", "set_cells"]
      and all(o.get("sheet") in ("AR_07.31", "Dashboard") or o["op"] == "duplicate_sheet" for o in p1),
      p1)
check("CL2: phase2 has exactly sales-paste+sales-setcells+raw-append",
      [o["sheet"] for o in p2] == ["New Sales report", "New Sales report", "Sales report Raw"],
      p2)
check("CL3: no op lost or duplicated", len(p1) + len(p2) == len(real_ops))

# ── correctness: two-phase result == single-phase result ──
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 5))
      + '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')
SHEETS = ["AR_06.30", "Dashboard", "New Sales report", "Sales report Raw"]
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheets>' + "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(SHEETS, 1))
      + '</sheets><calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, 5))
           + '</Relationships>')


def ws(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="A1:AH20"/><sheetData>{body}</sheetData></worksheet>')


ar = '<row r="6"><c r="A6" t="inlineStr"><is><t>hdr</t></is></c></row>'
dash = ''
sales = '<row r="6"><c r="A6" t="inlineStr"><is><t>hdr</t></is></c></row>'
raw = '<row r="3"><c r="A3" t="inlineStr"><is><t>hdr</t></is></c></row>'

WORK = tempfile.mkdtemp(prefix="twophase_")
src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    for i, body in enumerate([ar, dash, sales, raw], 1):
        z.writestr(f"xl/worksheets/sheet{i}.xml", ws(body))

ops = [
    {"op": "duplicate_sheet", "source": "AR_06.30", "name": "AR_07.31"},
    {"op": "set_cells", "sheet": "Dashboard", "cells": {"C2": "July'26"}},
    {"op": "paste_columns", "sheet": "AR_07.31", "anchor": "A7",
     "rows": [["cust1", 100.0], ["cust2", 200.0]]},
    {"op": "set_cells", "sheet": "AR_07.31", "cells": {"Z7": "=V7-L7", "Z8": "=V8-L8"}},
    {"op": "paste_columns", "sheet": "New Sales report", "anchor": "A7",
     "rows": [["s1", 50.0]]},
    {"op": "set_cells", "sheet": "New Sales report", "cells": {"Z7": "=Y7*2"}},
    {"op": "append_rows", "sheet": "Sales report Raw", "rows": [["r1", 1.0]]},
]


def run_ops(surgeon, op_list):
    for op in op_list:
        k = op["op"]
        if k == "set_cells":
            surgeon.set_cells(op["sheet"], op["cells"])
        elif k == "append_rows":
            surgeon.append_rows(op["sheet"], op["rows"])
        elif k == "duplicate_sheet":
            surgeon.duplicate_sheet(op["source"], op["name"])
        elif k == "paste_columns":
            surgeon.paste_columns(op["sheet"], op["anchor"], op["rows"],
                                  op.get("clear_beyond", True))


# single-phase
single_out = os.path.join(WORK, "single.xlsx")
s_single = XlsxSurgeon(src, workdir=WORK)
run_ops(s_single, ops)
res_single = s_single.apply(single_out)

# two-phase
p1, p2 = classify(ops, "AR_07.31")
mid = os.path.join(WORK, "mid.xlsx")
two_out = os.path.join(WORK, "two.xlsx")
s1 = XlsxSurgeon(src, workdir=WORK)
run_ops(s1, p1)
res1 = s1.apply(mid)
s2 = XlsxSurgeon(mid, workdir=WORK)
run_ops(s2, p2)
res2 = s2.apply(two_out)

with zipfile.ZipFile(single_out) as z:
    single_names = sorted(z.namelist())
    single_content = {n: z.read(n) for n in z.namelist() if n.endswith(".xml")}
with zipfile.ZipFile(two_out) as z:
    two_names = sorted(z.namelist())
    two_content = {n: z.read(n) for n in z.namelist() if n.endswith(".xml")}

# workbook.xml will differ trivially in sheet ordering/ids potentially, but
# worksheet DATA parts (the ones that matter) should be checkable by content
check("TP1: same set of part names", single_names == two_names,
      f"{single_names} vs {two_names}")

# find the AR_07.31 and New Sales report / Sales report Raw parts in each
# and confirm the actual row data matches regardless of which pass wrote it
import re as re_mod


def sheet_part(zip_path, sheet_name):
    with zipfile.ZipFile(zip_path) as z:
        wb = z.read("xl/workbook.xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
        rid = re_mod.search(rf'name="{re_mod.escape(sheet_name)}"[^>]*r:id="([^"]+)"', wb).group(1)
        tgt = re_mod.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1)
        return z.read("xl/" + tgt).decode()


single_ar = sheet_part(single_out, "AR_07.31")
two_ar = sheet_part(two_out, "AR_07.31")
check("TP2: AR_07.31 content identical single vs two-phase", single_ar == two_ar,
      f"single={single_ar[:200]}\ntwo={two_ar[:200]}")

single_sales = sheet_part(single_out, "New Sales report")
two_sales = sheet_part(two_out, "New Sales report")
check("TP3: New Sales report content identical", single_sales == two_sales)

single_raw = sheet_part(single_out, "Sales report Raw")
two_raw = sheet_part(two_out, "Sales report Raw")
check("TP4: Sales report Raw content identical", single_raw == two_raw)

single_dash = sheet_part(single_out, "Dashboard")
two_dash = sheet_part(two_out, "Dashboard")
check("TP5: Dashboard content identical", single_dash == two_dash)

check("TP6: two-phase opsResults non-empty for both phases",
      sum(r["cellsChanged"] for r in res1) > 0 and sum(r["cellsChanged"] for r in res2) > 0,
      (res1, res2))

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
