"""
Commission Workbook Surgery Service
-----------------------------------
Deploy on Render (Starter: 512MB RAM, disk-streamed where possible).

Security model: this service holds NO workbook credentials. n8n (which holds
the Microsoft OAuth) generates a pre-authenticated download URL
(@microsoft.graph.downloadUrl) and a pre-authenticated Graph upload session
URL, and passes both in the job request. URLs are short-lived Microsoft
tokens; the service just moves bytes.

The NetSuite extract runs HERE (not in n8n): the job's `extract` block makes
the service pull AR + sales rows from the NetSuite MCP and convert them into
surgeon ops. Requires env vars NS_MCP_URL and MCP_SHARED_SECRET.

Async job pattern (long HTTP requests get killed):
  POST /jobs                -> { jobId }         (returns immediately)
  GET  /jobs/{jobId}        -> { status: queued|running|done|failed, stage,
                                 arRows, salesRows, opsResults, ... }
n8n polls with the same Wait/poll loop used for the Graph copy monitor.

Job request body:
{
  "downloadUrl": "https://...  (pre-authed, from GET /items/{id} in n8n)",
  "uploadUrl":   "https://...  (pre-authed, from POST /createUploadSession)",
  "fileSize":    115343360,
  "stageOrder":  ["download","extract","edit","upload"],   # informational; this
                                                           # is already the fixed order
  "probe": [ {"sheet": "AR_06.30", "row": 7} ],  # optional: report the real
                                                 # formulas found on that row
  "extract": { "asofDate": "...", "fromDate": "...", "toDate": "...",
               "signFlip": true,
               "ar":    {"target": "AR_07.31", "anchor": "A7", "formulaCols": "Y:AH"},
               "sales": {"target": "New Sales report", "anchor": "A7", "formulaCols": "Z:AH"},
               "raw":   {"target": "Sales report Raw", "anchor": "A4", "formulaCols": "Z:AA"} },
  "ops": [
    {"op": "duplicate_sheet", "source": "AR_06.30", "name": "AR_07.31"},
    {"op": "set_cells",   "sheet": "Dashboard", "cells": {"C2": "July'26"}},
    {"op": "retarget_refs", "sheet": "AR_07.31",
     "replace": [{"from": "AR_05.31", "to": "AR_06.30"},
                 {"from": "AR_06.30", "to": "AR_07.31"}]}
  ]
}

retarget_refs: rolls the duplicated sheet's inherited cross-sheet
references forward one month. n8n sends BOTH mappings of the month roll:
two-months-back -> prior (for data-style refs) AND prior -> the new tab
itself (for header formulas like L2 = prior!V4 - self!L3). Mappings are
applied SIMULTANEOUSLY in one pass, so the chained rename cannot cascade,
and they touch only INHERITED content — formulas the extract writes are
excluded by construction. Both 'AR_05.31'! (quoted) and AR_05.31!
(unquoted) forms are matched. The job FAILS if ANY supplied mapping makes
zero replacements (reported per sheet as perMapping in opsResults).
Unknown top-level fields are REJECTED (422). This is deliberate: the model
used to ignore extras, and a payload whose `extract` block was silently
dropped reported "done" after uploading an unmodified 111MB file.

Ops generated from the extract are APPENDED after the caller's ops —
duplicate_sheet must run before the paste that targets the new sheet.

Cell values: numbers/strings as-is; strings starting with "=" are formulas.
"""

import json
import os
import re
import threading
import time
import traceback
import uuid
import zipfile

import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, ConfigDict

from surgeon import XlsxSurgeon
from commission_job import run_extract, build_ops
import xlsx_read as xr

try:
    import resource   # Linux/Render only — absent on Windows, diagnostic-only

    def _peak_rss_mb() -> float | None:
        return round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 1)
except ImportError:
    def _peak_rss_mb() -> float | None:
        return None


def _mem_checkpoint(j: dict, stage: str) -> None:
    """Record peak RSS so far at a pipeline stage boundary — both into the
    job status (visible while it's still running) and to stdout (Render's
    log stream survives an OOM kill even though the in-memory JOBS dict does
    not, so this is the one thing that lets a crash be diagnosed after the
    fact instead of guessed at again)."""
    mb = _peak_rss_mb()
    if mb is None:
        return
    checkpoints = j.setdefault("memCheckpointsMB", {})
    checkpoints[stage] = mb
    print(f"[mem] peak RSS after {stage}: {mb} MB", flush=True)

app = FastAPI(title="commission-workbook-surgeon")
JOBS: dict[str, dict] = {}
WORK = "/tmp/surgeon"
os.makedirs(WORK, exist_ok=True)

CHUNK_DL = 8 * 1024 * 1024
# Graph upload sessions require chunks in multiples of 320 KiB; 10 MiB works well.
CHUNK_UL = 32 * 320 * 1024        # 10,485,760 bytes

PROBE_SCAN_CAP = 64 * 1024 * 1024   # stop scanning a sheet part for the probe row

# ── job-state persistence ──────────────────────────────────────────────
# JOBS used to be memory-only: a Render redeploy/restart/crash mid-job made
# GET /jobs/{id} 404 ("unknown job") and the n8n poll loop errored out even
# when the job had actually finished. State now round-trips through a JSON
# file in /tmp — SURVIVES in-place process restarts (crash/OOM recovery
# reuses the same container: instance tags in Render logs persist across
# "Instance failed"/"Service recovered" cycles) but NOT redeploys, which
# swap in a fresh container and filesystem (verified live 2026-08-19: a
# redeploy mid-job 404'd the poll). Full redeploy survival would need
# Render's Persistent Disk or external state; crash recovery is the
# failure mode this service has actually hit.
# NOTE: this design (and the polling model) assumes WEB_CONCURRENCY=1.
# Multiple workers would each hold their own JOBS view and polls would land
# on the wrong process — if worker count ever needs raising, job state must
# move to real shared storage first.
JOBS_FILE = os.path.join(WORK, "jobs.json")
_JOBS_LOCK = threading.Lock()


def _persist_jobs() -> None:
    """Write-on-transition, atomically (temp file + os.replace) so a crash
    mid-write can never leave corrupt JSON that breaks the next startup.
    Volume is tiny — a handful of writes per job."""
    with _JOBS_LOCK:
        tmp = JOBS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(JOBS, f)
        os.replace(tmp, JOBS_FILE)


def _load_jobs() -> dict:
    try:
        with open(JOBS_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    now = time.time()
    out = {}
    for jid, j in data.items():
        if now - j.get("createdAt", now) > 7 * 86400:
            continue                     # prune entries older than ~7 days
        if j.get("status") in ("queued", "running"):
            # the thread that owned this job died with the old process —
            # mark it failed or n8n polls it forever
            j.update(status="failed",
                     error="service restarted while job was running "
                           f"(stage was {j.get('stage', 'unknown')!r})")
        out[jid] = j
    return out


JOBS.update(_load_jobs())
# sweep leaked work files: src+out+part files exceed 230MB per job and leak
# on a hard crash; anything older than an hour is dead weight
_sweep_now = time.time()
for _fn in os.listdir(WORK):
    _p = os.path.join(WORK, _fn)
    try:
        if (_fn != "jobs.json" and os.path.isfile(_p)
                and _sweep_now - os.path.getmtime(_p) > 3600):
            os.remove(_p)
    except OSError:
        pass
if JOBS:
    _persist_jobs()      # persist the restarted-while-running markings


class Job(BaseModel):
    model_config = ConfigDict(extra="forbid")
    downloadUrl: str
    uploadUrl: str
    fileSize: int | None = None
    extract: dict | None = None
    ops: list[dict] = []
    # n8n sends this; the service's order is fixed and already matches it.
    stageOrder: list[str] | None = None
    # optional formula probe: [{"sheet": "AR_06.30", "row": 7}, ...]
    probe: list[dict] | None = None
    # keep the output on disk and serve it at GET /jobs/{id}/result
    # (testing aid — the file survives until the next restart/redeploy)
    keepResult: bool = False
    # read cell values back to the caller AFTER ops are applied:
    # [{"sheet": "Dashboard", "range": "W6:W24", "as": "endingBalances"}]
    readRanges: list[dict] | None = None
    # fail the job on layout drift BEFORE pasting:
    # [{"sheet": "AR_07.31", "row": 6, "expect": "A:X"}]
    headerGuard: list[dict] | None = None
    # read-only workbook discovery (parts/pivots/rows); when omitted on an
    # extract job a default discovery spec runs — it is cheap and its output
    # feeds the reporting-half design work
    inspect: dict | None = None


# Verified header layout (ground truth: the AGA saved-search CSV export +
# the delivered workbook). Tuples = accepted variants. Comparison is
# whitespace-normalized and case-insensitive — the guard exists to catch
# COLUMN SHIFT, not cosmetic drift.
_AR_HDR = [
    "Client:Project", "Customer_Type", "Client Category: Name",
    "First Sale Date", "Store_Type", "Date", "Transaction Type", "No.",
    "Item: Full Name", "Item: Description (Sales)", "Quantity",
    "Open Balance", "Sales Rep: Name (Grouped)",
    "Address: Shipping Address State", "Partner",
    "Partner: Partner Category/Role", "Commission Pct",
    "Transaction Status: Description", "Due Date", "Date Closed",
    "Primary Partner: Name", "Amount (Gross)", "Account: Name (GL-style)",
    "Company Name",
]
_SALES_HDR = list(_AR_HDR[:22])
_SALES_HDR[11] = ("Amount", "Open Balance")     # L differs between tabs
_SALES_HDR[15] = "Client: Partner"              # P: person, not category
_SALES_HDR += ["Month", "Client Age (Years)", ("Company name", "Company Name")]
EXPECTED_HEADERS = {"A:X": _AR_HDR, "A:Y": _SALES_HDR}


def _hdr_norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().casefold()


def _run_header_guard(src: str, guards: list[dict], ops: list[dict]) -> list[dict]:
    """SOP step 7: verify the header row matches the verified layout before
    anything is pasted. Raises on the first mismatch, naming sheet, column,
    expected and actual. A guarded sheet that a duplicate_sheet op will
    CREATE is checked against its dup SOURCE's header (that is where the
    header will come from)."""
    dup_source = {op.get("name"): op.get("source")
                  for op in ops if op.get("op") == "duplicate_sheet"}
    s = XlsxSurgeon(src, workdir=WORK)
    report = []
    with zipfile.ZipFile(src) as zf:
        fetched = []
        for gd in guards or []:
            sheet, row = gd["sheet"], int(gd["row"])
            expect = EXPECTED_HEADERS.get(str(gd["expect"]).upper())
            if expect is None:
                raise ValueError(f"headerGuard: unknown expect {gd['expect']!r}")
            read_sheet = sheet if sheet in s._sheet_parts else dup_source.get(sheet)
            if not read_sheet or read_sheet not in s._sheet_parts:
                raise ValueError(f"headerGuard: sheet {sheet!r} not found "
                                 "(and no duplicate_sheet op creates it)")
            cells = xr.stream_rows(zf, s._sheet_parts[read_sheet],
                                   row, row).get(row, {})
            fetched.append((sheet, read_sheet, row, expect, cells))
        shared = xr.resolve_shared(
            zf, set().union(*[xr.shared_indices(c.values())
                              for *_x, c in fetched]) if fetched else set())
    for sheet, read_sheet, row, expect, cells in fetched:
        for i, exp in enumerate(expect):
            letter = xr.col_letter(i + 1)
            actual = xr.cell_value(cells.get(letter), shared)
            variants = exp if isinstance(exp, tuple) else (exp,)
            if _hdr_norm(actual) not in {_hdr_norm(v) for v in variants}:
                raise ValueError(
                    f"headerGuard FAILED on {sheet!r} (read from "
                    f"{read_sheet!r} row {row}) column {letter}: expected "
                    f"{variants[0]!r}, found {actual!r} — refusing to paste "
                    "into a shifted layout")
        report.append({"sheet": sheet, "readFrom": read_sheet, "row": row,
                       "columns": len(expect), "status": "ok"})
    return report


def _run_read_ranges(path: str, specs: list[dict], j: dict) -> None:
    """Read requested ranges from the FINISHED output (after ops), resolving
    shared strings, and expose them keyed by their 'as' names. Values are
    row-major lists of row-arrays (callers flatten). Sheets over the 32MB
    in-memory threshold are refused per-range, not per-job."""
    s = XlsxSurgeon(path, workdir=WORK)
    out: dict = {}
    errors: dict = {}
    with zipfile.ZipFile(path) as zf:
        pending = []
        for spec in specs or []:
            name = spec.get("as") or spec.get("range")
            sheet, rng = spec.get("sheet"), spec.get("range")
            part = s._sheet_parts.get(sheet)
            if not part:
                errors[name] = f"sheet {sheet!r} not found"
                continue
            if zf.getinfo(part).file_size > 32 * 1024 * 1024:
                errors[name] = (f"sheet {sheet!r} is "
                                f"{zf.getinfo(part).file_size // 1048576}MB "
                                "decompressed — refusing to read ranges off it")
                continue
            try:
                c_lo, c_hi, r_lo, r_hi = xr.parse_range(rng)
            except ValueError as e:
                errors[name] = str(e)
                continue
            rows = xr.stream_rows(zf, part, r_lo or 1, r_hi or 10 ** 7)
            pending.append((name, c_lo, c_hi, r_lo, r_hi, rows))
        need = set()
        for *_a, rows in pending:
            for rc in rows.values():
                need |= xr.shared_indices(rc.values())
        shared = xr.resolve_shared(zf, need)
    for name, c_lo, c_hi, r_lo, r_hi, rows in pending:
        if r_lo is not None:
            row_nums = list(range(r_lo, r_hi + 1))
        else:
            row_nums = sorted(rows)
        vals = []
        for rn in row_nums:
            rc = rows.get(rn, {})
            row_vals = [xr.cell_value(rc.get(xr.col_letter(c)), shared)
                        for c in range(c_lo, c_hi + 1)]
            if r_lo is None and all(v is None for v in row_vals):
                continue          # unbounded ranges skip fully-empty rows
            vals.append(row_vals)
        out[name] = vals
    if out:
        j["readRanges"] = out
    if errors:
        j["readRangesErrors"] = errors


_DISCOVERY_ROWS = [
    {"sheet": "Data", "rows": "1:12"},          # 1:12 reaches the data-area
    {"sheet": "Compiled Data", "rows": "1:10"},  # header (~row 9-10) + row 11
    {"sheet": "Payment", "rows": "1:40"},
    {"sheet": "Dashboard", "rows": "1:40"},
    {"sheet": "Info", "rows": "1:40"},
    {"sheet": "Kevin Hanks", "rows": "1:30"},
    {"sheet": "Tiffany M.", "rows": "1:30"},
    {"sheet": "Partial Payments", "rows": "1:12"},
    {"sheet": "Summary Pivot", "rows": "1:12"},
    {"sheet": "Statement Pivot", "rows": "1:8"},
    {"sheet": "Journal_Template", "rows": "1:20"},
    {"sheet": "REC", "rows": "1:10"},
    {"sheet": "Sales report", "rows": "1:6"},
    {"sheet": "Kevin Hanks_Sales", "rows": "1:6"},
    # current-month AR block inside Data (starts ~11+prior rows) and the
    # new-sales block start — the SOP's formulas-not-values exception rows
    {"sheet": "Data", "rows": "16368:16380"},
    {"sheet": "Data", "rows": "32720:32740"},
    # AR header-area parameter cells the rate formulas anchor to
    # ($AD$1, $AE$1, $AI$1, $AD$4)
    {"sheet": "AR_06.30", "rows": "1:6"},
    {"sheet": "Commission Rate by SKUs", "rows": "1:8"},
]


def _run_inspect(src: str, spec: dict) -> dict:
    """Read-only workbook discovery on the DOWNLOADED source copy: part
    inventory with sizes and dimensions, pivot-table definitions (the
    question that decides whether the reporting half is automatable at
    all), and value+formula dumps of requested head rows. Never fatal."""
    s = XlsxSurgeon(src, workdir=WORK)
    out: dict = {"sheetNames": s.sheet_names()}
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        if spec.get("parts", True):
            parts = {}
            for sheet, part in s._sheet_parts.items():
                try:
                    parts[sheet] = {
                        "part": part,
                        "bytes": zf.getinfo(part).file_size,
                        "dimension": xr.part_dimension(zf, part)}
                except (KeyError, OSError) as e:
                    parts[sheet] = {"part": part, "error": str(e)[:120]}
            out["parts"] = parts
            out["pivotParts"] = [n for n in names if "pivot" in n.lower()]
        if spec.get("pivots", True):
            pivots = []
            for pt in [n for n in names
                       if re.match(r"xl/pivotTables/pivotTable\d+\.xml$", n)]:
                try:
                    px = zf.read(pt).decode("utf-8", "replace")
                    info = {"part": pt}
                    m = re.search(r'<pivotTableDefinition[^>]*\bname="([^"]*)"', px)
                    info["name"] = m.group(1) if m else None
                    m = re.search(r'\bcacheId="(\d+)"', px)
                    info["cacheId"] = m.group(1) if m else None
                    rels = re.sub(r"(pivotTables/)(pivotTable\d+\.xml)$",
                                  r"\1_rels/\2.rels", pt)
                    cache_part = None
                    if rels in names:
                        rx = zf.read(rels).decode("utf-8", "replace")
                        cm = re.search(r'Target="([^"]*pivotCacheDefinition[^"]*)"', rx)
                        if cm:
                            cache_part = "xl/" + cm.group(1).replace("../", "")
                    field_names = []
                    if cache_part and cache_part in names:
                        cx = zf.read(cache_part).decode("utf-8", "replace")
                        sm = re.search(r'<worksheetSource[^>]*\bref="([^"]*)"[^>]*\bsheet="([^"]*)"', cx) \
                            or re.search(r'<worksheetSource[^>]*\bsheet="([^"]*)"[^>]*\bref="([^"]*)"', cx)
                        if sm:
                            g1, g2 = sm.group(1), sm.group(2)
                            ref, sheet = (g1, g2) if ":" in g1 or g1[:1].isalpha() and any(ch.isdigit() for ch in g1) else (g2, g1)
                            info["cacheSource"] = {"sheet": sheet, "ref": ref}
                        field_names = re.findall(r'<cacheField[^>]*\bname="([^"]*)"', cx)
                        info["cachePart"] = cache_part
                    def fname(ix):
                        try:
                            ix = int(ix)
                        except ValueError:
                            return ix
                        if ix == -2:
                            return "(values)"
                        return field_names[ix] if 0 <= ix < len(field_names) else f"field{ix}"
                    rf = re.search(r"<rowFields\b.*?</rowFields>", px, re.S)
                    cf = re.search(r"<colFields\b.*?</colFields>", px, re.S)
                    info["rowFields"] = [fname(x) for x in
                                         re.findall(r'<field x="(-?\d+)"', rf.group(0))] if rf else []
                    info["colFields"] = [fname(x) for x in
                                         re.findall(r'<field x="(-?\d+)"', cf.group(0))] if cf else []
                    info["dataFields"] = re.findall(r'<dataField[^>]*\bname="([^"]*)"', px)
                    pivots.append(info)
                except Exception as e:  # noqa: BLE001
                    pivots.append({"part": pt, "error": str(e)[:200]})
            out["pivots"] = pivots
        row_specs = spec.get("rows", _DISCOVERY_ROWS)
        dumped: dict = {}
        need = set()
        fetched = []
        for rs in row_specs:
            sheet = rs["sheet"]
            part = s._sheet_parts.get(sheet)
            if not part:
                dumped[sheet] = {"error": "sheet not found"}
                continue
            m = re.match(r"^(\d+):(\d+)$", str(rs["rows"]))
            lo, hi = (int(m.group(1)), int(m.group(2))) if m else (1, 8)
            rows = xr.stream_rows(zf, part, lo, hi)
            merged = next((f for f in fetched if f[0] == sheet), None)
            if merged:
                merged[1].update(rows)
            else:
                fetched.append((sheet, rows))
            for rc in rows.values():
                need |= xr.shared_indices(rc.values())
        shared = xr.resolve_shared(zf, need)
        for sheet, rows in fetched:
            sheet_dump = {}
            for rn in sorted(rows):
                rowd = {}
                for col, cell in rows[rn].items():
                    v = xr.cell_value(cell, shared)
                    if isinstance(v, str) and len(v) > 200:
                        v = v[:200] + "…"
                    ent = {}
                    if v is not None:
                        ent["v"] = v
                    if cell.get("f"):
                        ent["f"] = cell["f"][:300]
                    if ent:
                        rowd[col] = ent
                if rowd:
                    sheet_dump[str(rn)] = rowd
            dumped[sheet] = sheet_dump
        out["rows"] = dumped
    return out


def _auto_probes(job: "Job") -> list[dict]:
    """When an extract runs, probe the formula rows the templates must be
    read from: the dup-source AR sheet and the sales/raw target sheets at
    their anchor rows. Reported in job['formulaProbe'] so a real run hands
    back the workbook's actual formulas."""
    probes = {(p.get("sheet"), int(p.get("row", 0))) for p in (job.probe or [])}
    ext = job.extract or {}

    def row_of(block):
        m = re.search(r"(\d+)$", str((block or {}).get("anchor", "")))
        return int(m.group(1)) if m else None

    for op in job.ops:
        if op.get("op") == "duplicate_sheet" and row_of(ext.get("ar")):
            probes.add((op.get("source"), row_of(ext.get("ar"))))
    for key in ("sales", "raw"):
        blk = ext.get(key)
        if blk and blk.get("target") and row_of(blk):
            probes.add((blk["target"], row_of(blk)))
    return [{"sheet": s, "row": r} for s, r in sorted(probes) if s and r]


def _probe_formulas(src: str, probes: list[dict]) -> dict:
    """Report every formula cell on the requested rows, plus whether each
    probed sheet has Excel Table parts (structured refs on a duplicated copy
    would still resolve to the SOURCE sheet's table — report, don't guess).

    Streams only the head of each sheet part: target rows sit near the top,
    and 'Sales report Raw' is far too large to hold in memory."""
    s = XlsxSurgeon(src, workdir=WORK)
    # sheetNames = the SOURCE workbook's tab inventory, before any ops run.
    # Lets a stray tab (e.g. "AR_1.31 (2)") be attributed to the source
    # month's file vs something this job created, without opening 111MB.
    out: dict = {"formulas": {}, "tables": {}, "errors": [],
                 "sheetNames": s.sheet_names()}
    with zipfile.ZipFile(src) as zf:
        names = set(zf.namelist())
        for p in probes:
            sheet, row = p.get("sheet"), int(p.get("row", 0))
            part = s._sheet_parts.get(sheet)
            if not part or not row:
                out["errors"].append(f"probe {p!r}: sheet not found")
                continue
            # table parts hang off the sheet's rels
            rels = re.sub(r"(worksheets/)(sheet\d+\.xml)$", r"\1_rels/\2.rels", part)
            tables = []
            if rels in names:
                rx = zf.read(rels).decode("utf-8", "replace")
                tables = re.findall(r'Target="[^"]*tables/([^"]+)"', rx)
            out["tables"][sheet] = tables

            row_xml, buf, scanned = None, "", 0
            open_pat = f'<row r="{row}"'
            end_pat = "</row>"
            with zf.open(part) as f:
                while scanned < PROBE_SCAN_CAP:
                    chunk = f.read(4 * 1024 * 1024)
                    if not chunk:
                        break
                    scanned += len(chunk)
                    buf += chunk.decode("utf-8", "replace")
                    i = buf.find(open_pat)
                    if i == -1:
                        buf = buf[-200:]        # keep tag-boundary overlap only
                        continue
                    j = buf.find(end_pat, i)
                    if j == -1:
                        if buf.find("/>", i) == buf.find(">", i) - 1:
                            row_xml = buf[i:buf.find(">", i) + 1]
                            break
                        buf = buf[i:]
                        continue
                    row_xml = buf[i:j + len(end_pat)]
                    break
            if row_xml is None:
                out["errors"].append(f"probe {sheet} row {row}: row not found "
                                     f"in first {scanned // 1048576}MB")
                continue
            cells = {}
            for m in re.finditer(
                    r'<c r="([A-Z]+\d+)"[^>]*>\s*<f[^>]*>(.*?)</f>', row_xml, re.S):
                f = re.sub(r"\s+", " ", m.group(2)).strip()
                for ent, ch in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                                ("&apos;", "'"), ("&amp;", "&")):
                    f = f.replace(ent, ch)
                cells[m.group(1)] = "=" + f
            out["formulas"][f"{sheet}!r{row}"] = cells
    return out


def _upload_file(j: dict, path: str, upload_url: str) -> dict:
    """Chunked PUT to a Graph upload session using the documented RESUME
    protocol. The naive version re-PUT the same byte range after a client-
    side connection error — but if that PUT had actually LANDED server-side
    and it was the final chunk, Graph finalizes the file and destroys the
    session, so the blind retry hit a dead session and failed the whole
    ~15-minute job with 404 "The upload session was not found" (seen live
    twice). Now: after any connection failure, GET the session for
    nextExpectedRanges and continue from where the SERVER says it is; a
    404 on that status probe after the final chunk means the upload in
    fact completed — finish with a verification caveat instead of failing."""
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        pos = 0
        resp = None
        stalls = 0
        while pos < size:
            f.seek(pos)
            chunk = f.read(CHUNK_UL)
            end = pos + len(chunk) - 1
            try:
                resp = requests.put(
                    upload_url, data=chunk, timeout=600,
                    headers={"Content-Length": str(len(chunk)),
                             "Content-Range": f"bytes {pos}-{end}/{size}"})
            except (requests.exceptions.ConnectionError,
                    requests.exceptions.Timeout) as ue:
                stalls += 1
                if stalls > 6:
                    raise
                print(f"[ul] chunk {pos}-{end} broke ({ue}); asking the "
                      "session where it stands", flush=True)
                time.sleep(5)
                try:
                    st = requests.get(upload_url, timeout=60)
                except (requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout):
                    continue          # probe also failed; retry same chunk
                if st.status_code == 404:
                    if end + 1 >= size:
                        # final chunk landed before the connection dropped;
                        # Graph finalized and tore down the session
                        j["uploadedPct"] = 100
                        j["uploadCaveat"] = (
                            "final chunk connection dropped after the server "
                            "finalized the upload — n8n's Get Final File size "
                            "check is the confirmation")
                        print("[ul] session finalized during final-chunk "
                              "retry; treating as complete", flush=True)
                        return {}
                    raise RuntimeError(
                        "upload session disappeared mid-file (Graph 404 on "
                        "status probe) — the target file may have been "
                        "deleted from SharePoint during the run")
                nxt = (st.json().get("nextExpectedRanges") or [f"{pos}-"])[0]
                pos = int(str(nxt).split("-")[0])
                j["uploadedPct"] = round(pos / size * 100)
                continue
            if resp.status_code == 404:
                raise RuntimeError(
                    "upload chunk failed 404 (upload session not found) — "
                    "either the TESTP2 target file was deleted from "
                    "SharePoint mid-run, or the session expired. Re-run; do "
                    f"not clean the target folder while a run is in flight. "
                    f"Body: {resp.text[:200]}")
            if resp.status_code not in (200, 201, 202):
                raise RuntimeError(
                    f"upload chunk failed {resp.status_code}: {resp.text[:300]}")
            pos = end + 1
            j["uploadedPct"] = round(pos / size * 100)
        return resp.json() if (resp is not None and resp.text) else {}


def _run(job_id: str, job: Job):
    j = JOBS[job_id]
    src = os.path.join(WORK, f"{job_id}_src.xlsx")
    dst = os.path.join(WORK, f"{job_id}_out.xlsx")
    mid = os.path.join(WORK, f"{job_id}_mid.xlsx")   # phase-1 output for a 2-phase surgery
    try:
        j.update(status="running", stage="download")
        _persist_jobs()
        # a single unguarded GET on a ~116MB pull dies to any transient
        # reset (seen live: ConnectionResetError(104) mid-stream from the
        # Graph CDN). Retry with Range resumption — Graph download URLs
        # honor Range, so a retry continues from the break instead of
        # restarting the whole file.
        attempts = 0
        while True:
            attempts += 1
            try:
                pos = os.path.getsize(src) if os.path.exists(src) else 0
                hdrs = {"Range": f"bytes={pos}-"} if pos else {}
                with requests.get(job.downloadUrl, stream=True, timeout=600,
                                  headers=hdrs) as r:
                    if pos and r.status_code != 206:
                        pos = 0        # server ignored Range; restart clean
                    r.raise_for_status()
                    with open(src, "ab" if pos else "wb") as f:
                        for chunk in r.iter_content(CHUNK_DL):
                            f.write(chunk)
                break
            except (requests.exceptions.ConnectionError,
                    requests.exceptions.ChunkedEncodingError,
                    requests.exceptions.Timeout) as de:
                if attempts >= 4:
                    raise
                j["downloadRetries"] = attempts
                print(f"[dl] attempt {attempts} broke at "
                     f"{os.path.getsize(src) if os.path.exists(src) else 0} bytes "
                     f"({de}); retrying", flush=True)
                time.sleep(5 * attempts)
        if job.fileSize and os.path.getsize(src) != job.fileSize:
            raise ValueError(
                f"download incomplete: got {os.path.getsize(src)} bytes, "
                f"expected {job.fileSize} — refusing to operate on a truncated file")
        j["downloadedMB"] = round(os.path.getsize(src) / 1048576, 1)
        _mem_checkpoint(j, "download")
        _persist_jobs()

        # fail on layout drift BEFORE spending 5 minutes on the extract —
        # a guard failure here is the job working as designed
        if job.headerGuard:
            j["stage"] = "headerGuard"
            j["headerGuard"] = _run_header_guard(src, job.headerGuard, job.ops)
            _persist_jobs()

        if job.inspect is not None or job.extract:
            j["stage"] = "inspect"
            try:
                j["inspect"] = _run_inspect(src, job.inspect or {})
            except Exception as ie:  # noqa: BLE001 — discovery must never kill a job
                j["inspect"] = {"error": str(ie)[:400]}
            _persist_jobs()

        probes = _auto_probes(job) if job.extract else (job.probe or [])
        if probes:
            j["stage"] = "probe"
            try:
                j["formulaProbe"] = _probe_formulas(src, probes)
            except Exception as pe:  # noqa: BLE001 — the probe must never kill a job
                j["formulaProbe"] = {"error": str(pe)[:400]}
            _mem_checkpoint(j, "probe")
            _persist_jobs()

        if job.extract:
            j["stage"] = "extract"
            log_lines: list[str] = []

            def _log(msg):
                log_lines.append(str(msg))
                j["extractLog"] = log_lines[-30:]

            data = run_extract(job.extract, log=_log)
            # the AR formula templates' prior-month XLOOKUP needs the real
            # tab AR_07.31 (say) was duplicated FROM — derive it from the
            # caller's own duplicate_sheet op rather than requiring n8n to
            # send it separately.
            dup_op = next((o for o in job.ops
                          if o.get("op") == "duplicate_sheet"), None)
            prior_ar_tab = dup_op["source"] if dup_op else None
            gen_ops, report = build_ops(data, job.extract, prior_ar_tab=prior_ar_tab)
            j.update(
                arRows=data["arCount"], salesRows=data["salesCount"],
                arOpenBalance=data["arOpenBalance"], txnCount=data["txnCount"],
                extractDiagnostics=data["diagnostics"], opsReport=report,
                newItems=data.get("newItems", []),
                newItemsBasis=data.get("newItemsBasis"),
            )
            if data["arCount"] < 1000 or data["salesCount"] < 100:
                raise ValueError(
                    f"extract returned {data['arCount']} AR / {data['salesCount']} sales "
                    "rows — refusing to write. Expected ~16,000 / ~13,500.")
            if job.extract.get("applyOps", True):
                job.ops = list(job.ops) + gen_ops  # APPEND: duplicate_sheet must run first
            else:
                j["opsSkipped"] = "extract.applyOps=false — extract verified, no pastes generated"
            # build_ops's _pad() already made independent copies of every
            # row into gen_ops (now merged into job.ops) — data["arRows"]/
            # data["salesRows"] (16k+ and 13.5k+ Python lists) are a
            # redundant SECOND full copy that served no further purpose but
            # stayed resident for the rest of the job, padding the baseline
            # apply() started from (measured 185MB before phase 1 even
            # began — surprisingly high for "downloaded + extracted rows").
            del data, gen_ops
            _mem_checkpoint(j, "extract")
            _persist_jobs()

        j["stage"] = "surgery"
        _persist_jobs()
        # Split into two sequential apply() calls when an extract ran: AR's
        # duplicate+paste+formulas (~16k rows) and the sales sheet's own
        # paste+formulas (~13.5k rows) are each big enough to matter on
        # their own, and doing them in the SAME apply() call means their
        # transient overhead can be simultaneously resident even after each
        # op is logically done — CPython's allocator often doesn't return
        # freed memory to the OS within a long-running process, so RSS can
        # stay elevated from phase 1 while phase 2 is still running. Writing
        # phase 1's result to a local temp file and opening a BRAND NEW
        # XlsxSurgeon on it for phase 2 means phase 1's surgeon instance,
        # its queued ops, and its transform temporaries are all out of scope
        # before phase 2 begins — the two phases' peaks never overlap. The
        # extra round-trip is local disk I/O (fast), not network.
        ar_target = ((job.extract or {}).get("ar") or {}).get("target")
        phase1_ops, phase2_ops = list(job.ops), []
        if ar_target:
            phase1_ops, phase2_ops = [], []
            for op in job.ops:
                is_ar_or_dash = (
                    (op["op"] == "duplicate_sheet" and op.get("name") == ar_target)
                    or op.get("sheet") in (ar_target, "Dashboard"))
                (phase1_ops if is_ar_or_dash else phase2_ops).append(op)

        def _run_ops(surgeon, ops):
            for op in ops:
                kind = op["op"]
                if kind == "set_cells":
                    surgeon.set_cells(op["sheet"], op["cells"])
                elif kind == "append_rows":
                    surgeon.append_rows(op["sheet"], op["rows"])
                elif kind == "add_sheet":
                    surgeon.add_sheet(op["name"], op["rows"])
                elif kind == "duplicate_sheet":
                    surgeon.duplicate_sheet(op["source"], op["name"])
                elif kind == "paste_columns":
                    surgeon.paste_columns(op["sheet"], op["anchor"], op["rows"],
                                          op.get("clear_beyond", True))
                elif kind == "retarget_refs":
                    surgeon.retarget_refs(op["sheet"], op["replace"])
                elif kind == "replace_formula_text":
                    surgeon.replace_formula_text(op["sheet"], op["replace"])
                elif kind == "copy_range_values":
                    surgeon.copy_range_values(op["sheet"], op["from"], op["to"])
                elif kind == "pivot_refresh_on_load":
                    surgeon.pivot_refresh_on_load()
                else:
                    raise ValueError(f"unknown op {kind!r}")

        if phase2_ops:
            s1 = XlsxSurgeon(src, workdir=WORK)
            _run_ops(s1, phase1_ops)
            _mem_checkpoint(j, "before_apply_phase1")
            results1 = s1.apply(mid)
            _mem_checkpoint(j, "after_apply_phase1")
            # ensure phase 1's surgeon, its queued ops, AND the AR paste
            # rows/formula dicts (the bulk of the ~186MB pre-surgery
            # baseline) are all collectible before phase 2 begins
            del s1
            phase1_ops = None
            job.ops = []
            import gc
            gc.collect()
            _mem_checkpoint(j, "after_phase1_release")
            s2 = XlsxSurgeon(mid, workdir=WORK)   # fresh read from the intermediate
            _run_ops(s2, phase2_ops)
            _mem_checkpoint(j, "before_apply_phase2")
            results2 = s2.apply(dst)
            _mem_checkpoint(j, "after_apply_phase2")
            results = results1 + results2
        else:
            s = XlsxSurgeon(src, workdir=WORK)
            _run_ops(s, phase1_ops)
            _mem_checkpoint(j, "before_apply")
            results = s.apply(dst)
            _mem_checkpoint(j, "after_apply")
        j["opsResults"] = results
        j["outputMB"] = round(os.path.getsize(dst) / 1048576, 1)
        _persist_jobs()

        if job.readRanges:
            j["stage"] = "readRanges"
            try:
                _run_read_ranges(dst, job.readRanges, j)
            except Exception as re_err:  # noqa: BLE001 — reads must not kill the job
                j["readRangesErrors"] = {"_fatal": str(re_err)[:400]}
            # newItems exclusion: drop SKUs already on 'Commission Rate by
            # SKUs' when that range was requested (first column = SKU)
            sku_rows = (j.get("readRanges") or {}).get("skuRates")
            if sku_rows and j.get("newItems"):
                def _norm_sku(v):
                    if isinstance(v, float) and v.is_integer():
                        return str(int(v))
                    return str(v).strip()
                known = {_norm_sku(r[0]) for r in sku_rows if r and r[0] is not None}
                before = len(j["newItems"])
                j["newItems"] = [it for it in j["newItems"]
                                 if _norm_sku(it["sku"]) not in known]
                j["newItemsExcludedExisting"] = before - len(j["newItems"])
            _persist_jobs()

        j["stage"] = "upload"
        _persist_jobs()
        final = _upload_file(j, dst, job.uploadUrl)
        _mem_checkpoint(j, "upload")
        j.update(status="done", stage="complete",
                 resultItemId=final.get("id"),
                 resultWebUrl=final.get("webUrl"))
        _persist_jobs()
    except Exception as e:  # noqa: BLE001
        _mem_checkpoint(j, f"failed_at_{j.get('stage', 'unknown')}")
        j.update(status="failed", error=str(e)[:800],
                 trace=traceback.format_exc()[-1200:])
        _persist_jobs()
    finally:
        # keep the finished workbook when explicitly asked (keepResult), and
        # ALSO when the job died during upload — the surgery took ~10 min of
        # download+extract+transform and the output is complete; throwing it
        # away over an upload-session failure forces a full redo when
        # POST /jobs/{id}/retry-upload with a fresh uploadUrl would do.
        upload_failed = (j.get("status") == "failed"
                         and j.get("stage") == "upload")
        keep = ({dst} if ((job.keepResult or upload_failed)
                          and os.path.exists(dst)) else set())
        if keep:
            j["resultPath"] = dst
            if upload_failed:
                j["recovery"] = ("output kept — POST /jobs/{id}/retry-upload "
                                 "with a fresh uploadUrl to finish without "
                                 "redoing the extract/surgery")
        for p in (src, dst, mid):
            if p in keep:
                continue
            try:
                os.remove(p)
            except OSError:
                pass
        _persist_jobs()


VERSION = "2026-08-21-v23-dateroll"


@app.get("/health")
def health():
    return {"ok": True, "version": VERSION,
            "extractConfigured": bool(os.environ.get("NS_MCP_URL"))
            and bool(os.environ.get("MCP_SHARED_SECRET"))}


@app.post("/jobs")
def create_job(job: Job):
    job_id = uuid.uuid4().hex[:12]
    JOBS[job_id] = {"status": "queued", "createdAt": time.time()}
    _persist_jobs()
    threading.Thread(target=_run, args=(job_id, job), daemon=True).start()
    return {"jobId": job_id}


@app.get("/jobs")
def list_jobs():
    """Summaries only — lets an operator find the latest job id after the
    fact (e.g. to pull inspect/readRanges data from a run n8n started)."""
    return {jid: {"status": j.get("status"), "stage": j.get("stage"),
                  "createdAt": j.get("createdAt")}
            for jid, j in JOBS.items()}


@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(
            404, "unknown job — either it never existed, it aged out "
                 "(>7 days), or its state was lost to a REDEPLOY (a deploy "
                 "replaces the container and /tmp with it; only in-place "
                 "crash/OOM restarts preserve job state). Re-run the job.")
    return JOBS[job_id]


class UploadRetry(BaseModel):
    model_config = ConfigDict(extra="forbid")
    uploadUrl: str


@app.post("/jobs/{job_id}/retry-upload")
def retry_upload(job_id: str, body: UploadRetry):
    """Re-upload a kept result to a FRESH Graph upload session — recovery
    for jobs whose surgery finished but whose upload session died
    (itemNotFound: target file deleted mid-run, session expired, etc.).
    Mint a new uploadUrl (new server-side copy + createUploadSession in
    n8n) and pass it here; the finished workbook uploads in ~1 minute
    instead of redoing the whole extract+surgery."""
    j = JOBS.get(job_id)
    if not j:
        raise HTTPException(404, "unknown job")
    path = j.get("resultPath")
    if not path or not os.path.exists(path):
        raise HTTPException(409, "no kept result for this job (service "
                                 "restarted, or the job failed before the "
                                 "output was built) — re-run the full job")

    def _do():
        try:
            j.update(status="running", stage="upload", error=None)
            _persist_jobs()
            final = _upload_file(j, path, body.uploadUrl)
            j.update(status="done", stage="complete",
                     resultItemId=final.get("id"),
                     resultWebUrl=final.get("webUrl"))
        except Exception as e:  # noqa: BLE001
            j.update(status="failed", error=str(e)[:800],
                     trace=traceback.format_exc()[-1200:])
        _persist_jobs()

    threading.Thread(target=_do, daemon=True).start()
    return {"jobId": job_id, "status": "uploading"}


@app.get("/jobs/{job_id}/result")
def get_job_result(job_id: str):
    j = JOBS.get(job_id)
    if not j:
        raise HTTPException(404, "unknown job")
    path = j.get("resultPath")
    if not path or not os.path.exists(path):
        raise HTTPException(404, "no kept result for this job (keepResult not "
                                 "set, job not finished, or service restarted)")
    from fastapi.responses import FileResponse
    return FileResponse(path, filename=f"{job_id}_out.xlsx")
