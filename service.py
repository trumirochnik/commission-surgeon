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
    {"op": "set_cells",   "sheet": "Dashboard", "cells": {"C2": "July'26"}}
  ]
}
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
import traceback
import uuid
import zipfile

import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, ConfigDict

from surgeon import XlsxSurgeon
from commission_job import run_extract, build_ops

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


def _run(job_id: str, job: Job):
    j = JOBS[job_id]
    src = os.path.join(WORK, f"{job_id}_src.xlsx")
    dst = os.path.join(WORK, f"{job_id}_out.xlsx")
    try:
        j.update(status="running", stage="download")
        with requests.get(job.downloadUrl, stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(src, "wb") as f:
                for chunk in r.iter_content(CHUNK_DL):
                    f.write(chunk)
        j["downloadedMB"] = round(os.path.getsize(src) / 1048576, 1)

        probes = _auto_probes(job) if job.extract else (job.probe or [])
        if probes:
            j["stage"] = "probe"
            try:
                j["formulaProbe"] = _probe_formulas(src, probes)
            except Exception as pe:  # noqa: BLE001 — the probe must never kill a job
                j["formulaProbe"] = {"error": str(pe)[:400]}

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

        j["stage"] = "surgery"
        s = XlsxSurgeon(src, workdir=WORK)
        for op in job.ops:
            kind = op["op"]
            if kind == "set_cells":
                s.set_cells(op["sheet"], op["cells"])
            elif kind == "append_rows":
                s.append_rows(op["sheet"], op["rows"])
            elif kind == "add_sheet":
                s.add_sheet(op["name"], op["rows"])
            elif kind == "duplicate_sheet":
                s.duplicate_sheet(op["source"], op["name"])
            elif kind == "paste_columns":
                s.paste_columns(op["sheet"], op["anchor"], op["rows"],
                                op.get("clear_beyond", True))
            else:
                raise ValueError(f"unknown op {kind!r}")
        results = s.apply(dst)      # raises if no op changed a single cell
        j["opsResults"] = results
        j["outputMB"] = round(os.path.getsize(dst) / 1048576, 1)

        j["stage"] = "upload"
        size = os.path.getsize(dst)
        with open(dst, "rb") as f:
            pos = 0
            while pos < size:
                chunk = f.read(CHUNK_UL)
                end = pos + len(chunk) - 1
                resp = requests.put(
                    job.uploadUrl, data=chunk, timeout=600,
                    headers={"Content-Length": str(len(chunk)),
                             "Content-Range": f"bytes {pos}-{end}/{size}"})
                if resp.status_code not in (200, 201, 202):
                    raise RuntimeError(
                        f"upload chunk failed {resp.status_code}: {resp.text[:300]}")
                pos = end + 1
                j["uploadedPct"] = round(pos / size * 100)
            final = resp.json() if resp.text else {}
        j.update(status="done", stage="complete",
                 resultItemId=final.get("id"),
                 resultWebUrl=final.get("webUrl"))
    except Exception as e:  # noqa: BLE001
        j.update(status="failed", error=str(e)[:800],
                 trace=traceback.format_exc()[-1200:])
    finally:
        keep = {dst} if (job.keepResult and os.path.exists(dst)) else set()
        if keep:
            j["resultPath"] = dst
        for p in (src, dst):
            if p in keep:
                continue
            try:
                os.remove(p)
            except OSError:
                pass


VERSION = "2026-08-17-extract-v6-memhardening"


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
