"""
Commission Workbook Surgery Service
-----------------------------------
Deploy on Render (free tier compatible: disk-streamed, <300MB RAM).

Security model: this service holds NO credentials. n8n (which holds the
Microsoft OAuth) generates a pre-authenticated download URL
(@microsoft.graph.downloadUrl) and a pre-authenticated Graph upload session
URL, and passes both in the job request. URLs are short-lived Microsoft
tokens; the service just moves bytes.

Async job pattern (Render free tier kills long HTTP requests):
  POST /jobs                -> { jobId }         (returns immediately)
  GET  /jobs/{jobId}        -> { status: queued|running|done|failed, detail }
n8n polls with the same Wait/poll loop used for the Graph copy monitor.

Job request body:
{
  "downloadUrl": "https://...  (pre-authed, from GET /items/{id} in n8n)",
  "uploadUrl":   "https://...  (pre-authed, from POST /createUploadSession)",
  "fileSize":    115343360,
  "ops": [
    {"op": "set_cells",   "sheet": "Dashboard", "cells": {"B2": "July'26", "B3": "June'26"}},
    {"op": "append_rows", "sheet": "Data",      "rows": [[...], ...]},
    {"op": "add_sheet",   "name": "AR_07.31",   "rows": [[...], ...]}
  ]
}
Cell values: numbers/strings as-is; strings starting with "=" are formulas.
"""

import os
import threading
import traceback
import uuid

import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

from surgeon import XlsxSurgeon

app = FastAPI(title="commission-workbook-surgeon")
JOBS: dict[str, dict] = {}
WORK = "/tmp/surgeon"
os.makedirs(WORK, exist_ok=True)

CHUNK_DL = 8 * 1024 * 1024
CHUNK_UL = 10 * 320 * 1024 * 32   # 100MB? no — see below
# Graph upload sessions require chunks in multiples of 320 KiB; 10 MiB works well.
CHUNK_UL = 32 * 320 * 1024        # 10,485,760 bytes


class Job(BaseModel):
    downloadUrl: str
    uploadUrl: str
    fileSize: int | None = None
    ops: list[dict]


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

        j["stage"] = "surgery"
        s = XlsxSurgeon(src, workdir=WORK)
        applied = []
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
            applied.append(kind)
        s.apply(dst)
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
                 opsApplied=applied,
                 resultItemId=final.get("id"),
                 resultWebUrl=final.get("webUrl"))
    except Exception as e:  # noqa: BLE001
        j.update(status="failed", error=str(e)[:800],
                 trace=traceback.format_exc()[-1200:])
    finally:
        for p in (src, dst):
            try:
                os.remove(p)
            except OSError:
                pass


@app.get("/health")
def health():
    return {"ok": True}


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
