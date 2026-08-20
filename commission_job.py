"""
commission_job.py — wires netsuite_extract into the existing XlsxSurgeon.

Drop this next to service.py and surgeon.py. It deliberately does NOT touch
surgeon.py: it only produces ops that the existing dispatch already supports
(duplicate_sheet, paste_columns, append_rows, set_cells), so no writer changes
are required to get the extract flowing.

────────────────────────────────────────────────────────────────────────
WHY THE OPS ARE SHAPED THE WAY THEY ARE
────────────────────────────────────────────────────────────────────────

paste_columns writes a rectangular block whose width is max(len(row)). Cells
inside the block are overwritten; cells to the RIGHT of the block survive on
rows that already exist. New rows past the old last row are created containing
ONLY the block's columns — nothing is copied down.

That gives one hard constraint: the pasted block must be exactly the DATA
columns and no wider. If it were widened to cover the formula columns, every
formula cell in Y:AH would be overwritten with None (which deletes it) on all
16k existing rows, destroying the sheet. So:

  * data goes in via paste_columns at the data width (24 for AR, 25 for sales)
  * formula columns are written separately, as text starting with '=' — the
    surgeon turns those into real formulas

Sheet size guard: paste_columns/set_cells need the target part < 32MB
decompressed. AR_07.31 is exempt because duplicate_sheet applies it in memory
in the same job. 'Sales report Raw' is cumulative (row 464,908 as of 06.2026)
and far over — hence append_rows, which has no limit and streams.
"""

from __future__ import annotations

import os
import re
from typing import Any

from netsuite_extract import (
    FORMULA_TEMPLATES,
    Mcp,
    col_to_index,
    extract as ns_extract,
)

AR_DATA_COLS = 24        # A:X
SALES_DATA_COLS = 25     # A:Y


def _pad(rows: list[list], width: int) -> list[list]:
    """Every row exactly `width` long. Guards against a short row silently
    narrowing the whole pasted block."""
    out = []
    for r in rows:
        if len(r) > width:
            raise ValueError(f"row is {len(r)} wide, expected at most {width}")
        out.append(list(r) + [None] * (width - len(r)))
    return out


def _formula_cells(kind: str, first_row: int, count: int,
                   prior_ar_tab: str | None = None) -> dict[str, str]:
    """{'Z7': '=V7-L7', ...} for the rows we just wrote.

    Only templates present in FORMULA_TEMPLATES are emitted. Columns with no
    template are left alone — on existing rows the duplicated sheet's own
    formulas survive; on rows past the old extent they stay EMPTY. That is a
    real gap, reported by build_ops() as `formulaGaps`.
    """
    from netsuite_extract import formula_cells as _fc
    return _fc(kind, first_row, count, prior_ar_tab=prior_ar_tab)


def _span_cols(span: str) -> list[str]:
    m = re.match(r"^([A-Z]+):([A-Z]+)$", span.strip().upper())
    if not m:
        return []
    lo, hi = col_to_index(m.group(1)), col_to_index(m.group(2))
    out = []
    for n in range(lo, hi + 1):
        s, x = "", n
        while x:
            x, rem = divmod(x - 1, 26)
            s = chr(65 + rem) + s
        out.append(s)
    return out


def _anchor_row(anchor: str) -> int:
    m = re.search(r"(\d+)$", anchor)
    if not m:
        raise ValueError(f"bad anchor {anchor!r}")
    return int(m.group(1))


def build_ops(data: dict, spec: dict, prior_ar_tab: str | None = None
             ) -> tuple[list[dict], dict]:
    """extract results + the payload's `extract` spec -> ops for the surgeon.

    `spec` is the job's extract block: {ar:{target,anchor,formulaCols,writeMode},
    sales:{...}, raw:{...}}. `prior_ar_tab` is the sheet AR_07.31 (say) was
    duplicated FROM (e.g. 'AR_06.30') — two AR formula columns (AD, AG) do a
    prior-month XLOOKUP and need the real tab name, not last month's hardcoded
    one carried forward. Returns (ops, report).
    """
    ops: list[dict] = []
    report: dict[str, Any] = {"formulaGaps": {}, "rowCounts": {}}

    ar_rows = data["arRows"]
    sales_rows = data["salesRows"]

    # ---- AR: duplicated sheet, so paste_columns is exempt from the 32MB guard
    ar = spec["ar"]
    ar_first = _anchor_row(ar["anchor"])
    ops.append({"op": "paste_columns", "sheet": ar["target"],
                "anchor": ar["anchor"], "rows": _pad(ar_rows, AR_DATA_COLS),
                "clear_beyond": True})
    cells = _formula_cells("ar", ar_first, len(ar_rows), prior_ar_tab=prior_ar_tab)
    if cells:
        ops.append({"op": "set_cells", "sheet": ar["target"], "cells": cells})
    have = set(re.match(r"^([A-Z]+)", k).group(1) for k in cells)
    report["formulaGaps"]["ar"] = [c for c in _span_cols(ar.get("formulaCols", ""))
                                   if c not in have]
    report["rowCounts"]["ar"] = len(ar_rows)

    # ---- New Sales report: EXISTING sheet -> 32MB guard applies
    sl = spec["sales"]
    sl_first = _anchor_row(sl["anchor"])
    ops.append({"op": "paste_columns", "sheet": sl["target"],
                "anchor": sl["anchor"], "rows": _pad(sales_rows, SALES_DATA_COLS),
                "clear_beyond": True})
    cells = _formula_cells("sales", sl_first, len(sales_rows))
    if cells:
        ops.append({"op": "set_cells", "sheet": sl["target"], "cells": cells})
    have = set(re.match(r"^([A-Z]+)", k).group(1) for k in cells)
    report["formulaGaps"]["sales"] = [c for c in _span_cols(sl.get("formulaCols", ""))
                                      if c not in have]
    report["rowCounts"]["sales"] = len(sales_rows)

    # ---- Sales report Raw: cumulative, append only
    raw = spec.get("raw")
    if raw:
        width = SALES_DATA_COLS
        tpl = FORMULA_TEMPLATES.get("raw") or {}
        if tpl:
            # appended rows are brand new, so widening the block is safe here
            width = max(col_to_index(c) for c in tpl)
        padded = _pad(sales_rows, width)
        if tpl:
            # {r} is deliberately NOT substituted here: append_rows doesn't
            # know a row's final sheet position until the surgeon actually
            # writes it (it's appended past whatever the sheet's current
            # last row is). row_xml() fills in the real row number at that
            # point — see surgeon.py.
            for row in padded:
                for col, t in tpl.items():
                    row[col_to_index(col) - 1] = t
        ops.append({"op": "append_rows", "sheet": raw["target"], "rows": padded})
        report["formulaGaps"]["raw"] = [c for c in _span_cols(raw.get("formulaCols", ""))
                                       if c not in tpl]
        report["rowCounts"]["raw"] = len(sales_rows)

    return ops, report


def build_prior_ops(rows: list[list], spec: dict, asof_serial: int
                    ) -> tuple[list[dict], dict]:
    """Ops for the prior-tab refresh (see extract_prior_ar): re-paste the
    prior month's AR data (24 cols, current Date Closed in T) and convert
    the tab to the PRIOR LAYOUT — Y/Z/AD/AE/AF/AG formulas plus AA/AB/AC
    values, and the $AB$4 future-month gate in the header (a row-4 cell,
    which the streaming rebuild writes into the prefix). Rows past the new
    extent are dropped by the stream — including the hand-paste's shifted
    junk rows whose stale T values spuriously passed the receipt cutoff."""
    from netsuite_extract import prior_formula_cells, prior_value_cells
    target = spec["target"]
    anchor = spec.get("anchor", "A7")
    first = _anchor_row(anchor)
    cells: dict[str, object] = {}
    cells.update(prior_formula_cells(first, len(rows)))
    cells.update(prior_value_cells(rows, first, asof_serial))
    for ref, val in (spec.get("headerCells") or {}).items():
        cells[ref] = val
    ops = [
        {"op": "paste_columns", "sheet": target, "anchor": anchor,
         "rows": _pad(rows, AR_DATA_COLS), "clear_beyond": True},
        {"op": "set_cells", "sheet": target, "cells": cells},
    ]
    return ops, {"rowCount": len(rows), "cellCount": len(cells)}


def run_extract(spec: dict, log=print) -> dict:
    """Run the NetSuite extract described by the job's `extract` block."""
    url = spec.get("mcpUrl") or os.environ.get("NS_MCP_URL")
    secret = os.environ.get("MCP_SHARED_SECRET")
    if not url:
        raise RuntimeError("NS_MCP_URL not set and no mcpUrl in the job")
    if not secret:
        raise RuntimeError("MCP_SHARED_SECRET not set on the service")
    mcp = Mcp(url, secret)
    return ns_extract(mcp, spec["asofDate"], spec["fromDate"], spec["toDate"],
                      sign_flip=spec.get("signFlip", True), log=log)


def run_prior_extract(spec: dict, prior_spec: dict, log=print) -> dict:
    """AR-only pull for the prior-tab refresh. `spec` is the job's extract
    block (creds/flags), `prior_spec` its priorAr sub-block."""
    from netsuite_extract import extract_prior_ar
    url = spec.get("mcpUrl") or os.environ.get("NS_MCP_URL")
    secret = os.environ.get("MCP_SHARED_SECRET")
    if not url:
        raise RuntimeError("NS_MCP_URL not set and no mcpUrl in the job")
    if not secret:
        raise RuntimeError("MCP_SHARED_SECRET not set on the service")
    mcp = Mcp(url, secret)
    return extract_prior_ar(mcp, prior_spec["asofDate"],
                            sign_flip=spec.get("signFlip", True), log=log)
