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
     "replace": [{"from": "AR_05.31", "to": "AR_06.30"}]}
  ]
}

retarget_refs: rewrites cross-sheet references in one sheet's formulas —
the duplicated AR tab inherits the SOURCE month's prior-month references
(July's tab copied from June still points at May), so n8n should send
from = two-months-back tab name, to = the duplicate_sheet source. Both
'AR_05.31'! (quoted) and AR_05.31! (unquoted) forms are matched. The job
FAILS if a supplied mapping makes zero replacements.
Unknown top-level fields are REJECTED (422). This is deliberate: the model
used to ignore extras, and a payload whose `extract` block was silently
dropped reported "done" after uploading an unmodified 111MB file.

Ops generated from the extract are APPENDED after the caller's ops —
duplicate_sheet must run before the paste that targets the new sheet.

Cell values: numbers/strings as-is; strings starting with "=" are formulas.
"""

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
    out: dict = {"formulas": {}, "tables": {}, "errors": []}
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
    """Chunked PUT of `path` to a Graph upload session, with per-chunk retry
    on transient resets. Returns the final driveItem JSON (or {})."""
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        pos = 0
        resp = None
        while pos < size:
            chunk = f.read(CHUNK_UL)
            end = pos + len(chunk) - 1
            for attempt in range(1, 4):
                try:
                    resp = requests.put(
                        upload_url, data=chunk, timeout=600,
                        headers={"Content-Length": str(len(chunk)),
                                 "Content-Range": f"bytes {pos}-{end}/{size}"})
                    break
                except (requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout) as ue:
                    # Graph upload sessions are resumable; re-PUTting the
                    # same byte range after a transient reset is safe
                    if attempt >= 3:
                        raise
                    print(f"[ul] chunk {pos}-{end} attempt {attempt} "
                         f"broke ({ue}); retrying", flush=True)
                    time.sleep(5 * attempt)
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

        probes = _auto_probes(job) if job.extract else (job.probe or [])
        if probes:
            j["stage"] = "probe"
            try:
                j["formulaProbe"] = _probe_formulas(src, probes)
            except Exception as pe:  # noqa: BLE001 — the probe must never kill a job
                j["formulaProbe"] = {"error": str(pe)[:400]}
            _mem_checkpoint(j, "probe")

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

        j["stage"] = "surgery"
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

        j["stage"] = "upload"
        final = _upload_file(j, dst, job.uploadUrl)
        _mem_checkpoint(j, "upload")
        j.update(status="done", stage="complete",
                 resultItemId=final.get("id"),
                 resultWebUrl=final.get("webUrl"))
    except Exception as e:  # noqa: BLE001
        _mem_checkpoint(j, f"failed_at_{j.get('stage', 'unknown')}")
        j.update(status="failed", error=str(e)[:800],
                 trace=traceback.format_exc()[-1200:])
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


VERSION = "2026-08-18-extract-v13-uploadrecovery"


@app.get("/health")
def health():
    return {"ok": True, "version": VERSION,
            "extractConfigured": bool(os.environ.get("NS_MCP_URL"))
            and bool(os.environ.get("MCP_SHARED_SECRET"))}


@app.post("/jobs")
def create_job(job: Job):
    job_id = uuid.uuid4().hex[:12]
    JOBS[job_id] = {"status": "queued"}
    threading.Thread(target=_run, args=(job_id, job), daemon=True).start()
    return {"jobId": job_id}


@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, "unknown job")
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
            final = _upload_file(j, path, body.uploadUrl)
            j.update(status="done", stage="complete",
                     resultItemId=final.get("id"),
                     resultWebUrl=final.get("webUrl"))
        except Exception as e:  # noqa: BLE001
            j.update(status="failed", error=str(e)[:800],
                     trace=traceback.format_exc()[-1200:])

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
