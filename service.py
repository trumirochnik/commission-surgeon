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
from commission_job import run_extract, build_ops, build_prior_ops, run_prior_extract
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
    # OPTIONAL since v24: when omitted, the job stops at stage "edited" with
    # the output kept on disk, and the caller delivers it via
    # POST /jobs/{id}/retry-upload with a FRESHLY minted uploadUrl. Rationale:
    # a session minted before the job starts sits idle through ~15 min of
    # download+extract+surgery, and SharePoint reaps idle sessions — the
    # third distinct upload-404 incident (2026-08-19) died exactly there.
    # Late minting makes the session seconds old when the first byte flows.
    uploadUrl: str | None = None
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
    # per-rep statement generation from an ALREADY-RECALCULATED workbook:
    # {"periodLabel": "07.2026"} — downloads the file, reads Compiled Data
    # and Payment cached values, builds one small .xlsx per partner, and
    # keeps them zipped at resultPath (GET /jobs/{id}/result). Refused if
    # the workbook has not been opened in desktop Excel since surgery.
    statements: dict | None = None


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


def _read_prior_rows(path: str, sheet: str, first_row: int) -> list[list]:
    """The prior tab's data rows (A:X, 24 cols) exactly as pasted — the
    as-of balances Mike closed the month on are ground truth for the
    refresh; only Date Closed gets overwritten (see enrich_prior_rows).
    Streams the part and resolves only the shared strings those rows use."""
    s = XlsxSurgeon(path, workdir=WORK)
    part = s._sheet_parts.get(sheet)
    if not part:
        raise ValueError(f"prior tab {sheet!r} not found in the workbook")
    cols = [xr.col_letter(i) for i in range(1, 25)]
    with zipfile.ZipFile(path) as zf:
        rows = xr.stream_rows(zf, part, first_row, 10 ** 7)
        need = set()
        for rc in rows.values():
            need |= xr.shared_indices(rc.values())
        shared = xr.resolve_shared(zf, need)
    out = []
    for rn in sorted(rows):
        rc = rows[rn]
        vals = [xr.cell_value(rc.get(c), shared) for c in cols]
        if any(v is not None and v != "" for v in vals):
            out.append(vals)
    return out


def _read_sheet_cols(path: str, sheet: str, cols: list[str],
                     first_row: int, last_row: int = 10 ** 7) -> dict[int, dict]:
    """{row: {col: value}} for the requested columns, shared strings
    resolved. For the reporting reads (Compiled Data, Payment, SKU/Kevin
    rate tabs) — all comfortably small parts."""
    s = XlsxSurgeon(path, workdir=WORK)
    part = s._sheet_parts.get(sheet)
    if not part:
        raise ValueError(f"sheet {sheet!r} not found")
    with zipfile.ZipFile(path) as zf:
        rows = xr.stream_rows(zf, part, first_row, last_row)
        need = set()
        for rc in rows.values():
            need |= xr.shared_indices(rc.values())
        shared = xr.resolve_shared(zf, need)
    out: dict[int, dict] = {}
    for rn, rc in rows.items():
        vals = {c: xr.cell_value(rc.get(c), shared) for c in cols}
        if any(v is not None and v != "" for v in vals.values()):
            out[rn] = vals
    return out


def _run_reporting_phase(job_id: str, job: Job, data_spec: dict,
                         src: str, dst: str, j: dict) -> list[dict]:
    """Data tab (3 blocks) + Compiled Data range extension + new
    company/partner/rate combo rows. Runs AFTER phase 2 on its own surgeon
    so the ~46k-row ops never coexist with the paste phases' memory."""
    import pickle
    import report_data as rd
    from netsuite_extract import serial as _serial

    p_main = os.path.join(WORK, f"{job_id}_report_main.pkl")
    p_prior = os.path.join(WORK, f"{job_id}_report_prior.pkl")
    with open(p_main, "rb") as f:
        main = pickle.load(f)
    with open(p_prior, "rb") as f:
        prior_rows = pickle.load(f)

    prior_tab = job.extract["priorAr"]["target"]
    cur_tab = job.extract["ar"]["target"]
    asof_serial = _serial(job.extract["asofDate"])

    # rate inputs (read from the SOURCE workbook): the SKU rate table, the
    # licensed-item set (presence in its column B), and the header constants
    # the rate formulas reference
    sku = _read_sheet_cols(src, "Commission Rate by SKUs", ["B", "E"], 2)
    sku_rates = {str(v["B"]).strip(): v["E"] for v in sku.values()
                 if v.get("B") is not None and isinstance(v.get("E"), (int, float))}
    licensed_ids = {str(v["B"]).strip() for v in sku.values()
                    if v.get("B") is not None}
    consts = _shadow_consts(src, prior_tab)
    combos, undetermined = rd.distinct_combos(
        prior_rows, main["ar"], main["sales"], sku_rates, licensed_ids, consts)

    # SHADOW CALC + per-rep statements, in the same run. The engine is the
    # workbook's arithmetic ported 1:1 and penny-matched against the
    # reconciled 06.2026 book (all 18 partners, diff 0.00). Statements are
    # independent of the Dashboard's credit-memo adjustments — those touch
    # the accrual (V4/J), not the Data->Compiled chain.
    shadow = rd.shadow_compiled(prior_rows, main["ar"], main["sales"],
                                asof_serial, sku_rates, licensed_ids, consts)
    payment = rd.shadow_payment(shadow, _read_fee_table(src))
    period = data_spec.get("periodLabel") or data_spec.get("monthTag", "period")
    stmt_files = rd.build_statements(shadow, payment, period)
    zpath = os.path.join(WORK, f"{job_id}_statements.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in sorted(stmt_files.items()):
            z.writestr(name, blob)
        z.writestr("manifest.json", json.dumps(
            {"period": period, "count": len(stmt_files),
             "files": sorted(stmt_files)}, indent=1))
    j.update(statementsPath=zpath, statementCount=len(stmt_files),
             statementFiles=sorted(stmt_files),
             shadowEarnedByPartner={p: round(v["earned"], 2)
                                    for p, v in sorted(payment.items())
                                    if abs(v["earned"]) > 0.005})
    del shadow, payment, stmt_files

    built = rd.build_data_rows(prior_rows, main["ar"], main["sales"],
                               asof_serial, prior_tab, cur_tab,
                               "New Sales report", data_spec["earnedLabel"])
    del main, prior_rows

    # Compiled Data: current extent + existing combo list
    cd = _read_sheet_cols(src, "Compiled Data", ["C", "E", "F", "G"], 5)
    old_end = None
    existing = set()
    last_used = rd.CD_FIRST_DATA_ROW - 1
    for rn, v in sorted(cd.items()):
        f_g = None
        if v.get("C") not in (None, ""):
            last_used = max(last_used, rn)
            if v.get("E") not in (None, "") and isinstance(v.get("F"), (int, float)):
                existing.add((str(v["C"]), str(v["E"]), round(float(v["F"]), 6)))
    with zipfile.ZipFile(src) as zf:
        s0 = XlsxSurgeon(src, workdir=WORK)
        cxml = zf.read(s0._sheet_parts["Compiled Data"]).decode()
    m = re.search(r"Data!\$N\$11:\$N\$(\d+)", cxml)
    if not m:
        raise ValueError("Compiled Data: could not detect the Data range end")
    old_end = int(m.group(1))
    del cxml

    # EVERY bucket the data actually produces gets a row — new pairs AND
    # rate variants of existing pairs. The engine deciding this is
    # shadow_rate, which penny-matched the reconciled June book (all 18
    # partners, diff 0.00), so a variant here is a real bucket; skipping
    # them (the first, conservative policy) dropped those amounts out of
    # Compiled and broke the Summary Pivot tie-outs by exactly that much
    # (measured 11.9k on the 0821-1548 run).
    new_combos = sorted(c for c in combos if c not in existing)
    # both pivot caches read Compiled C4:T3723 — appended rows must stay inside
    room = 3700 - last_used
    dropped_combos = max(0, len(new_combos) - max(room, 0))
    new_combos = new_combos[:max(room, 0)]

    mid2 = os.path.join(WORK, f"{job_id}_mid2.xlsx")
    os.replace(dst, mid2)
    s3 = XlsxSurgeon(mid2, workdir=WORK)
    s3.paste_columns("Data", f"A{rd.DATA_FIRST_ROW}", built["pasteRows"],
                     clear_beyond=True)
    s3.set_cells("Data", dict(built["headerCells"]))
    if old_end != built["lastRow"]:
        s3.replace_formula_text("Compiled Data", [
            {"from": f"${old_end}", "to": f"${built['lastRow']}"}])
    if new_combos:
        s3.set_cells("Compiled Data", rd.compiled_combo_rows(
            new_combos, last_used + 1, built["lastRow"],
            data_spec.get("monthTag", "")))
    del built["pasteRows"]
    results = s3.apply(dst)
    del s3
    for p in (p_main, p_prior, mid2):
        try:
            os.remove(p)
        except OSError:
            pass
    j.update(dataLastRow=built["lastRow"], dataBlocks=built["blocks"],
             compiledRangeEnd=built["lastRow"],
             compiledCombosAdded=len(new_combos),
             compiledCombosDropped=dropped_combos,
             compiledCombosUndetermined=sorted(undetermined)[:25])
    return results


def _shadow_consts(path: str, prior_tab: str) -> dict:
    """Header constants the rate formulas reference: AD1 (Dayna date gate),
    AE1 (her name), AI1 (Bomgaar/Kelly date gate) on the AR tabs, AJ4 on
    the sales tab. Read from the workbook; documented defaults otherwise."""
    out = {"dayna": "Dayna Stambeck", "ad1": 45703, "ai1": 46053, "aj4": 45762}
    try:
        hdr = _read_sheet_cols(path, prior_tab, ["AD", "AE", "AI"], 1, 1)
        v = hdr.get(1, {})
        if isinstance(v.get("AD"), (int, float)):
            out["ad1"] = v["AD"]
        if isinstance(v.get("AE"), str) and v["AE"].strip():
            out["dayna"] = v["AE"].strip()
        if isinstance(v.get("AI"), (int, float)):
            out["ai1"] = v["AI"]
        s4 = _read_sheet_cols(path, "New Sales report", ["AJ"], 4, 4)
        if isinstance(s4.get(4, {}).get("AJ"), (int, float)):
            out["aj4"] = s4[4]["AJ"]
    except Exception:  # noqa: BLE001 — defaults are the documented values
        pass
    return out


def _read_fee_table(path: str) -> dict:
    """Payment N7:O24 — the static per-partner technology fees."""
    rows = _read_sheet_cols(path, "Payment", ["N", "O"], 7, 24)
    return {str(v["N"]).strip(): v["O"] for v in rows.values()
            if v.get("N") not in (None, "") and isinstance(v.get("O"), (int, float))}


def _build_statements_zip(job_id: str, src: str, spec: dict, j: dict) -> str:
    """Per-rep statement files from a RECALCULATED workbook's cached values."""
    import report_data as rd
    with zipfile.ZipFile(src) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode()
        if 'fullCalcOnLoad="1"' in wb_xml:
            raise ValueError(
                "this workbook has not been opened/recalculated in desktop "
                "Excel since surgery — Compiled Data still carries stale "
                "caches. Open it once, let the recalc finish, save, then "
                "re-run statements.")
    cd = _read_sheet_cols(src, "Compiled Data",
                          ["C", "E", "F", "G", "H", "I", "J", "K", "L", "S"],
                          rd.CD_FIRST_DATA_ROW)
    compiled = []
    for rn, v in sorted(cd.items()):
        if v.get("C") in (None, "") or v.get("E") in (None, ""):
            continue
        num = lambda x: x if isinstance(x, (int, float)) else 0.0
        compiled.append({
            "company": v["C"], "partner": v["E"], "rate": v.get("F"),
            "prior": num(v.get("G")), "newSales": num(v.get("H")),
            "collections": num(v.get("I")) + num(v.get("J")),
            "partial": num(v.get("K")), "totalColl": num(v.get("L")),
            "earned": num(v.get("S")),
        })
    pay_rows = _read_sheet_cols(src, "Payment", ["E", "F", "G", "H", "I"], 6, 25)
    payment = {}
    for rn, v in pay_rows.items():
        if v.get("E") in (None, "", "Grand Total"):
            continue
        num = lambda x: x if isinstance(x, (int, float)) else 0.0
        payment[str(v["E"])] = {"earned": num(v.get("F")), "fee": num(v.get("G")),
                                "adj": num(v.get("H")), "net": num(v.get("I"))}
    files = rd.build_statements(compiled, payment,
                                spec.get("periodLabel", "period"))
    if not files:
        raise ValueError("no partner had reportable activity — refusing to "
                         "emit an empty statements zip")
    zpath = os.path.join(WORK, f"{job_id}_statements.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in sorted(files.items()):
            z.writestr(name, blob)
        z.writestr("manifest.json", json.dumps(
            {"period": spec.get("periodLabel"), "count": len(files),
             "files": sorted(files)}, indent=1))
    j.update(statementCount=len(files), statementFiles=sorted(files))
    return zpath


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

        if job.statements:
            # lean job: no surgery — read the recalced workbook, emit the
            # per-rep statement files (SOP 'Reporting': one file per rep,
            # distribution stays with Jennifer)
            j["stage"] = "statements"
            _persist_jobs()
            zpath = _build_statements_zip(job_id, src, job.statements, j)
            j.update(status="done", stage="complete", resultPath=zpath,
                     recovery="GET /jobs/{id}/result returns the zip")
            _persist_jobs()
            return

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
                # SPILL the generated ops to disk, partitioned by the pass
                # that will run them. Holding every sheet's paste rows in
                # job.ops from extract to the end of surgery kept ~46k rows
                # resident (~350MB) and two runs OOM'd at the raw stage on
                # top of that plateau (0820 evening + 0821 morning). Each
                # pass now loads only its own rows.
                import pickle
                _ar_t = (job.extract.get("ar") or {}).get("target")
                spill = {"p1": [], "sales": [], "raw": []}
                for op in gen_ops:
                    if op["op"] == "append_rows":
                        spill["raw"].append(op)
                    elif op.get("sheet") == _ar_t:
                        spill["p1"].append(op)
                    else:
                        spill["sales"].append(op)
                for name, lst in spill.items():
                    if lst:
                        with open(os.path.join(WORK, f"{job_id}_ops_{name}.pkl"),
                                  "wb") as pf:
                            pickle.dump(lst, pf)
                del spill
                j["opsSpilled"] = True
            else:
                j["opsSkipped"] = "extract.applyOps=false — extract verified, no pastes generated"
            # prior-tab refresh (SOP step proven from the hand-built 06.2026
            # file): re-pull the PRIOR month's AR as-of its month end — run
            # now, so Date Closed carries this month's payments — and
            # convert that tab to the prior layout. Without it the receipts
            # term of Dashboard E is blind (measured: J 594k vs Mike's 0).
            prior_spec = job.extract.get("priorAr")
            if prior_spec and job.extract.get("applyOps", True):
                from netsuite_extract import serial as _serial, enrich_prior_rows
                close_map = run_prior_extract(job.extract, prior_spec, log=_log)
                if len(close_map) < 100:
                    raise ValueError(
                        f"prior-AR refresh: only {len(close_map)} close dates "
                        "came back — refusing to blank the tab's Date Closed "
                        "column against that.")
                first = int(re.search(r"(\d+)$",
                                      prior_spec.get("anchor", "A7")).group(1))
                prows = _read_prior_rows(src, prior_spec["target"], first)
                if len(prows) < 1000:
                    raise ValueError(
                        f"prior tab read back only {len(prows)} rows — "
                        "refusing to rewrite it.")
                hit = enrich_prior_rows(prows, close_map)
                bal = round(sum(r[11] for r in prows
                                if isinstance(r[11], (int, float))), 2)
                pops, preport = build_prior_ops(
                    prows, prior_spec, _serial(prior_spec["asofDate"]))
                import pickle
                with open(os.path.join(WORK, f"{job_id}_ops_prior.pkl"),
                          "wb") as pf:
                    pickle.dump(pops, pf)
                j["opsSpilled"] = True
                j.update(priorArRows=len(prows),
                         priorArOpenBalance=bal,
                         priorArClosedCount=hit,
                         priorArReport=preport)
                if job.extract.get("dataTab"):
                    import pickle
                    with open(os.path.join(WORK, f"{job_id}_report_prior.pkl"),
                              "wb") as pf:
                        pickle.dump(prows, pf)
                del close_map, prows, pops
            # reporting half: stash the row-sets on DISK for phase 3 — the
            # Data-tab ops (~46k rows + formula cells) are built only when
            # that phase runs, so phases 1-2 never hold them in memory
            data_spec = job.extract.get("dataTab")
            if data_spec and job.extract.get("applyOps", True):
                if not job.extract.get("priorAr"):
                    raise ValueError("extract.dataTab requires extract.priorAr "
                                     "— the Data tab's first block IS the "
                                     "refreshed prior AR")
                import pickle
                with open(os.path.join(WORK, f"{job_id}_report_main.pkl"),
                          "wb") as pf:
                    pickle.dump({"ar": data["arRows"],
                                 "sales": data["salesRows"]}, pf)
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

        if j.get("opsSpilled"):
            # ONE SHEET PER PASS, ops loaded lazily from the spill files —
            # each pass holds only its own rows (~40-80MB) instead of the
            # ~350MB every-sheet plateau that OOM'd two runs at the raw
            # stage. Caller ops are partitioned by target; pivot stamping
            # and anything unrecognized ride the final pass.
            import gc
            import pickle
            sales_target = ((job.extract or {}).get("sales") or {}).get("target")
            caller_p1, caller_sales, caller_last = [], [], []
            for op in job.ops:
                if op["op"] == "pivot_refresh_on_load":
                    caller_last.append(op)
                elif ((op["op"] == "duplicate_sheet" and op.get("name") == ar_target)
                        or op.get("sheet") in (ar_target, "Dashboard")):
                    caller_p1.append(op)
                elif op.get("sheet") == sales_target:
                    caller_sales.append(op)
                else:
                    caller_last.append(op)
            job.ops = []

            def _spill_path(name):
                return os.path.join(WORK, f"{job_id}_ops_{name}.pkl")

            def _load_spill(name):
                p = _spill_path(name)
                if not os.path.exists(p):
                    return []
                with open(p, "rb") as f:
                    lst = pickle.load(f)
                os.remove(p)
                return lst

            PASSES = [
                ("ar_dashboard", lambda: caller_p1 + _load_spill("p1")),
                ("sales", lambda: caller_sales + _load_spill("sales")),
                ("prior_refresh", lambda: _load_spill("prior")),
                ("raw_final", lambda: _load_spill("raw") + caller_last),
            ]
            has = {
                "ar_dashboard": bool(caller_p1) or os.path.exists(_spill_path("p1")),
                "sales": bool(caller_sales) or os.path.exists(_spill_path("sales")),
                "prior_refresh": os.path.exists(_spill_path("prior")),
                "raw_final": bool(caller_last) or os.path.exists(_spill_path("raw")),
            }
            labels = [lb for lb, _fn in PASSES if has[lb]]
            results = []
            cur_in = src
            for idx, (label, fn) in enumerate(p for p in PASSES if has[p[0]]):
                out = dst if idx == len(labels) - 1 else \
                    os.path.join(WORK, f"{job_id}_pass{idx}.xlsx")
                lst = fn()
                sgn = XlsxSurgeon(cur_in, workdir=WORK)
                _run_ops(sgn, lst)
                _mem_checkpoint(j, f"before_{label}")
                results += sgn.apply(out)
                _mem_checkpoint(j, f"after_{label}")
                del sgn, lst
                gc.collect()
                if cur_in != src:
                    try:
                        os.remove(cur_in)
                    except OSError:
                        pass
                cur_in = out
            caller_p1 = caller_sales = caller_last = None
            gc.collect()
        else:
            s = XlsxSurgeon(src, workdir=WORK)
            _run_ops(s, list(job.ops))
            _mem_checkpoint(j, "before_apply")
            results = s.apply(dst)
            _mem_checkpoint(j, "after_apply")
            del s
            job.ops = []
            import gc
            gc.collect()
        # phase 3: the reporting half — Data tab regeneration + Compiled
        # Data maintenance, on the phase-2 output with a fresh surgeon
        data_spec = (job.extract or {}).get("dataTab") if job.extract else None
        if data_spec and (job.extract or {}).get("applyOps", True):
            j["stage"] = "reporting"
            _persist_jobs()
            results += _run_reporting_phase(job_id, job, data_spec, src, dst, j)
            _mem_checkpoint(j, "reporting")

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

        if job.uploadUrl:
            j["stage"] = "upload"
            _persist_jobs()
            final = _upload_file(j, dst, job.uploadUrl)
            _mem_checkpoint(j, "upload")
            j.update(status="done", stage="complete",
                     resultItemId=final.get("id"),
                     resultWebUrl=final.get("webUrl"))
        else:
            # late-mint flow: surgery is finished, delivery happens via
            # retry-upload with a session minted moments before use
            j.update(status="done", stage="edited", awaitingUpload=True,
                     recovery="surgery complete — POST /jobs/{id}/retry-upload "
                              "with a fresh uploadUrl to deliver the result")
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
        awaiting = bool(j.get("awaitingUpload"))
        j["keepResult"] = bool(job.keepResult)   # retry-upload cleanup honors it
        keep = ({dst} if ((job.keepResult or upload_failed or awaiting)
                          and os.path.exists(dst)) else set())
        if keep:
            j["resultPath"] = dst
            if upload_failed:
                j["recovery"] = ("output kept — POST /jobs/{id}/retry-upload "
                                 "with a fresh uploadUrl to finish without "
                                 "redoing the extract/surgery")
        # spill/pass/report leftovers from a failed run
        import glob as _glob
        leftovers = (_glob.glob(os.path.join(WORK, f"{job_id}_ops_*.pkl"))
                     + _glob.glob(os.path.join(WORK, f"{job_id}_pass*.xlsx"))
                     + _glob.glob(os.path.join(WORK, f"{job_id}_report_*.pkl"))
                     + _glob.glob(os.path.join(WORK, f"{job_id}_mid2.xlsx")))
        for p in leftovers + [src, dst, mid]:
            if p in keep:
                continue
            try:
                os.remove(p)
            except OSError:
                pass
        _persist_jobs()


VERSION = "2026-08-21-v38-allbuckets"


@app.get("/health")
def health():
    return {"ok": True, "version": VERSION,
            "extractConfigured": bool(os.environ.get("NS_MCP_URL"))
            and bool(os.environ.get("MCP_SHARED_SECRET"))}


@app.post("/jobs")
def create_job(job: Job):
    # ONE job at a time. Two ~115MB jobs landed 83s apart on 2026-08-19
    # (two n8n workflows answered the same trigger) and ran concurrently —
    # combined RSS blew the 512MB container and BOTH died. A duplicate
    # trigger now gets a clear 409 instead of killing the run in flight.
    # Safe across crashes: restart marks running jobs failed at load time.
    for running_id, rj in JOBS.items():
        if rj.get("status") in ("queued", "running"):
            raise HTTPException(
                409, f"another job is already in flight (id {running_id}, "
                     f"stage {rj.get('stage', '?')!r}, started "
                     f"{int(time.time() - rj.get('createdAt', time.time()))}s ago) — "
                     "this container fits ONE workbook job. If you didn't "
                     "start it, two n8n workflows are answering the same "
                     "trigger: deactivate every Commission workflow except "
                     "the current one. A stale 'running' job clears on "
                     "service restart.")
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
            j.pop("awaitingUpload", None)
            # delivered — drop the 90MB+ temp copy unless the job asked to
            # keep it (a failed retry keeps resultPath so it can be retried)
            if not j.get("keepResult"):
                try:
                    os.remove(path)
                except OSError:
                    pass
                j.pop("resultPath", None)
        except Exception as e:  # noqa: BLE001
            j.update(status="failed", error=str(e)[:800],
                     trace=traceback.format_exc()[-1200:])
        _persist_jobs()

    threading.Thread(target=_do, daemon=True).start()
    return {"jobId": job_id, "status": "uploading"}


@app.get("/jobs/{job_id}/statements")
def get_job_statements(job_id: str):
    j = JOBS.get(job_id)
    if not j:
        raise HTTPException(404, "unknown job")
    path = j.get("statementsPath")
    if not path or not os.path.exists(path):
        raise HTTPException(404, "no statements for this job (dataTab not "
                                 "requested, job not finished, or service "
                                 "restarted)")
    from fastapi.responses import FileResponse
    return FileResponse(path, filename=f"{job_id}_statements.zip")


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
