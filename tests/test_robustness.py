"""FOLLOW-UP 3 robustness tests: job-state persistence semantics, no-op
double-run refusal, opsResults op/sheet keys."""
import importlib, json, os, re, sys, time, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
os.environ.setdefault("NS_MCP_URL", "https://example.invalid/mcp")
import service
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:250]}"))
    if not cond:
        fails.append(label)


# ── persistence: atomic write, load, restart-marking, pruning ──
TD = tempfile.mkdtemp(prefix="jobs_persist_")
service.JOBS_FILE = os.path.join(TD, "jobs.json")
service.JOBS.clear()
now = time.time()
service.JOBS.update({
    "runningjob": {"status": "running", "stage": "extract", "createdAt": now - 60},
    "queuedjob": {"status": "queued", "createdAt": now - 10},
    "donejob": {"status": "done", "stage": "complete", "createdAt": now - 3600,
                "opsResults": [{"op": "transform", "sheet": "X", "cellsChanged": 5}]},
    "ancient": {"status": "done", "createdAt": now - 8 * 86400},
})
service._persist_jobs()
check("P1: jobs.json written", os.path.exists(service.JOBS_FILE))
check("P2: no stray temp file left", not os.path.exists(service.JOBS_FILE + ".tmp"))

loaded = service._load_jobs()
check("P3: running job marked failed on load",
      loaded["runningjob"]["status"] == "failed"
      and "restarted while job was running" in loaded["runningjob"]["error"]
      and "extract" in loaded["runningjob"]["error"], loaded.get("runningjob"))
check("P4: queued job also marked failed", loaded["queuedjob"]["status"] == "failed")
check("P5: done job survives intact",
      loaded["donejob"]["status"] == "done"
      and loaded["donejob"]["opsResults"][0]["cellsChanged"] == 5)
check("P6: >7-day-old entry pruned", "ancient" not in loaded)

# corrupt file must not break startup
with open(service.JOBS_FILE, "w") as f:
    f.write('{"broken": tru')
check("P7: corrupt jobs.json -> empty load, no crash", service._load_jobs() == {})
service.JOBS.clear()

# ── no-op double-run + op/sheet keys, via the surgeon on a real zip ──
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
DASH = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1"/><sheetData/></worksheet>')

src = os.path.join(TD, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", DASH)

# run 1: writes C2 -> must count 1, must carry op/sheet keys
out1 = os.path.join(TD, "out1.xlsx")
s = XlsxSurgeon(src, workdir=TD)
s.set_cells("Dashboard", {"C2": "TESTVALUE"})
res1 = s.apply(out1)
check("N1: first write counts 1 cell",
      sum(r["cellsChanged"] for r in res1) == 1, res1)
check("N2: entries carry op/sheet/cellsChanged keys (n8n gate contract)",
      all("op" in r and "sheet" in r and "cellsChanged" in r for r in res1), res1)
check("N3: transform entry names the sheet",
      any(r["op"] == "transform" and r["sheet"] == "Dashboard" for r in res1), res1)

# run 2: SAME value onto run 1's output -> zero changes -> refuse
s2 = XlsxSurgeon(out1, workdir=TD)
s2.set_cells("Dashboard", {"C2": "TESTVALUE"})
try:
    s2.apply(os.path.join(TD, "out2.xlsx"))
    check("N4: no-op double-run refuses to produce output", False)
except ValueError as e:
    check("N4: no-op double-run refuses to produce output",
          "refusing to upload" in str(e), str(e))

# run 3: different value -> counts again
s3 = XlsxSurgeon(out1, workdir=TD)
s3.set_cells("Dashboard", {"C2": "OTHER"})
res3 = s3.apply(os.path.join(TD, "out3.xlsx"))
check("N5: changed value counts again", sum(r["cellsChanged"] for r in res3) == 1)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
