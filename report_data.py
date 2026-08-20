"""
report_data.py — the REPORTING half: Data tab, Compiled Data maintenance,
and per-rep statement files.

Reverse-engineered from the reconciled hand-built 06.2026 workbook
(2026-08-20). The chain is:

    Data (3 stacked blocks)  ->  Compiled Data (SUMIFS by company/partner/
    rate)  ->  Summary Pivot + Statement Pivot (caches read Compiled
    C4:T3723 / C4:T8724, already oversized)  ->  Payment (XLOOKUPs off
    Summary Pivot + static tech-fee table)  ->  per-rep statement files.

DATA TAB LAYOUT (row 10 = header, data from row 11):
  A Date Break (empty) | B Period label — EXACTLY 'Prior Month' /
  'Current Month' / 'New Sales' (Compiled's criteria match these against
  its own header cells G$4/H$4/I$3/J$3/M$4) | C..Z = the source tab's 24
  data columns A..X | AA Difference (CURRENT block only, as a FORMULA
  =X-N per the SOP exception) | AB Deposit Month (DATE serial; future
  months REMOVED per SOP — the AB>0 criterion is what counts a row as
  collected) | AC Client Age | AD Deposit Month number | AE Commission
  Rate | AF Unearned | AG Earned | AI =CONCAT(Z,J,K) lookup key |
  AJ prior-open-balance XLOOKUP (ranges pinned to the prior block's
  extent) | AK Amount (Gross) - Original (current block: VALUE copy) |
  AL Amount (Gross) - Adjusted (formula) | AM =CONCAT(J," - ",K).
  The SOP's "paste adjusted over Amount (Gross)" maneuver is expressed as
  Data X{r} = =AL{r} on the current block (X is not an AL input — no cycle).

BLOCK SOURCES (all row-aligned 1:1 with their tab, so AE/AF/AG can be
tab REFERENCES that recalc to the same values Mike pastes):
  Prior Month   <- the refreshed prior AR tab (prior layout: AD rate,
                   AE unearned, AF earned); AB/AC/AD are server-computed
                   VALUES (gated deposit date from the verified close map)
  Current Month <- the new AR tab (June layout: AA rate, AB unearned)
  New Sales     <- 'New Sales report' (AB dep date, X age, AC dep num,
                   AD rate, AE unearned, AF earned)
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile

EXCEL_EPOCH = dt.date(1899, 12, 30)

DATA_FIRST_ROW = 11
SPACER = 2                     # blank rows between blocks, as in the source

# Data C..Z <- source cols A..X (0-based indices 0..23) for AR-shaped rows.
# Sales rows: C..X <- sales A..V (22 cols), Y stays empty, Z <- sales Y
# (company, index 24) — verified against Data rows 30684+ in the June book.


def _serial_month(t) -> int:
    if not isinstance(t, (int, float)):
        return 0
    return (EXCEL_EPOCH + dt.timedelta(days=int(t))).month


def build_data_rows(prior_rows: list[list], cur_rows: list[list],
                    sales_rows: list[list], asof_serial: int,
                    prior_tab: str, cur_tab: str, sales_tab: str,
                    earned_label: str) -> dict:
    """-> {'pasteRows': [...26-wide A..Z...], 'cells': {ref: value|formula},
          'lastRow': n, 'blocks': {...}, 'headerCells': {...}}

    pasteRows carry A..Z (A empty, B period label, C..Z data). Everything
    from AA rightward goes through `cells` so formulas and per-block
    divergence stay explicit."""
    paste: list[list] = []
    cells: dict[str, object] = {}
    rn = DATA_FIRST_ROW

    def is_num(v):
        return isinstance(v, (int, float))

    # ---- block 1: Prior Month ------------------------------------------
    blk1_first = rn
    for i, row in enumerate(prior_rows):
        src = 7 + i                      # tab data starts at row 7
        paste.append([None, "Prior Month"] + list(row[:24]))
        t = row[19]                      # verified Date Closed serial
        if is_num(t) and t <= asof_serial:
            cells[f"AB{rn}"] = t         # SOP: future deposit months removed
            cells[f"AD{rn}"] = _serial_month(t)
        else:
            cells[f"AD{rn}"] = 0
        fs = row[3]
        if is_num(fs):
            cells[f"AC{rn}"] = round((asof_serial - fs) / 365, 2)
        cells[f"AE{rn}"] = f"='{prior_tab}'!AD{src}"    # rate
        cells[f"AF{rn}"] = f"='{prior_tab}'!AE{src}"    # unearned
        cells[f"AG{rn}"] = f"='{prior_tab}'!AF{src}"    # earned
        cells[f"AI{rn}"] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        cells[f"AM{rn}"] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        rn += 1
    blk1_last = rn - 1
    for _ in range(SPACER):              # paste is consecutive from the
        paste.append([None] * 26)        # anchor — spacers must be real rows
        rn += 1

    # ---- block 2: Current Month ----------------------------------------
    blk2_first = rn
    aj_rng = f"$AI${blk1_first}:$AI${blk1_last}"
    aj_n = f"$N${blk1_first}:$N${blk1_last}"
    for i, row in enumerate(cur_rows):
        src = 7 + i
        paste.append([None, "Current Month"] + list(row[:24]))
        cells[f"AA{rn}"] = f"=X{rn}-N{rn}"              # SOP: formula, not value
        cells[f"AE{rn}"] = f"='{cur_tab}'!AA{src}"      # rate (June layout)
        cells[f"AF{rn}"] = f"='{cur_tab}'!AB{src}"      # unearned
        cells[f"AI{rn}"] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        cells[f"AJ{rn}"] = (f'=IF($B{rn}="Current Month",_xlfn.XLOOKUP('
                            f'AI{rn},{aj_rng},{aj_n},"Error",0),0)')
        gross = row[21]                                  # V Amount (Gross)
        if gross is not None:
            cells[f"AK{rn}"] = gross                     # Original, as VALUE
        cells[f"AL{rn}"] = (f"=IF(AJ{rn}<0,IF(AJ{rn}>AK{rn},AJ{rn},"
                            f"IF(AJ{rn}<AK{rn},AJ{rn},AK{rn})),"
                            f"IF(AJ{rn}<AK{rn},AJ{rn},AK{rn}))")
        cells[f"X{rn}"] = f"=AL{rn}"    # the paste-adjusted-over-gross step
        cells[f"AM{rn}"] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        rn += 1
    blk2_last = rn - 1
    for _ in range(SPACER):
        paste.append([None] * 26)
        rn += 1

    # ---- block 3: New Sales --------------------------------------------
    blk3_first = rn
    for i, row in enumerate(sales_rows):
        src = 7 + i
        r25 = list(row[:25]) + [None] * (25 - len(row[:25]))
        paste.append([None, "New Sales"] + r25[:22] + [None, r25[24]])
        cells[f"AB{rn}"] = f"='{sales_tab}'!AB{src}"    # deposit date (gated)
        cells[f"AC{rn}"] = f"='{sales_tab}'!X{src}"     # client age
        cells[f"AD{rn}"] = f"='{sales_tab}'!AC{src}"    # deposit month number
        cells[f"AE{rn}"] = f"='{sales_tab}'!AD{src}"    # rate
        cells[f"AF{rn}"] = f"='{sales_tab}'!AE{src}"    # unearned
        cells[f"AG{rn}"] = f"='{sales_tab}'!AF{src}"    # earned
        cells[f"AI{rn}"] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        cells[f"AM{rn}"] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        rn += 1
    blk3_last = rn - 1

    return {
        "pasteRows": paste, "cells": cells, "lastRow": blk3_last,
        "blocks": {"prior": (blk1_first, blk1_last),
                   "current": (blk2_first, blk2_last),
                   "sales": (blk3_first, blk3_last)},
        # header zone (rows < 11 -> written into the streaming prefix)
        "headerCells": {"AD4": asof_serial, "AB8": asof_serial + 1,
                        "AG10": earned_label},
    }


# ── Compiled Data maintenance ─────────────────────────────────────────

CD_FIRST_DATA_ROW = 5
# row-5 formula skeleton, {r}-substituted for appended combo rows
_CD_ROW_TPL = {
    "G": ("=SUMIFS(Data!$N$11:$N${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$B$11:$B${end},'Compiled Data'!G$4,Data!$W$11:$W${end},"
          "'Compiled Data'!$E{r},Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "H": ("=SUMIFS(Data!$N$11:$N${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$B$11:$B${end},'Compiled Data'!H$4,Data!$W$11:$W${end},"
          "'Compiled Data'!$E{r},Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "I": ("=-SUMIFS(Data!$N$11:$N${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$B$11:$B${end},'Compiled Data'!I$3,Data!$W$11:$W${end},"
          "'Compiled Data'!$E{r},Data!$AB$11:$AB${end},\">0\","
          "Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "J": ("=-SUMIFS(Data!$N$11:$N${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$B$11:$B${end},'Compiled Data'!J$3,Data!$W$11:$W${end},"
          "'Compiled Data'!$E{r},Data!$AB$11:$AB${end},\">0\","
          "Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "K": ("=-SUMIFS(Data!$AA$11:$AA${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$W$11:$W${end},'Compiled Data'!$E{r},"
          "Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "L": "=SUM(I{r}:K{r})",
    "M": ("=SUMIFS(Data!$N$11:$N${end},Data!$Z$11:$Z${end},'Compiled Data'!$C{r},"
          "Data!$B$11:$B${end},'Compiled Data'!M$4,Data!$W$11:$W${end},"
          "'Compiled Data'!$E{r},Data!$AE$11:$AE${end},'Compiled Data'!$F{r})"),
    "N": "=G{r}+H{r}+L{r}",
    "O": "=M{r}-N{r}",
    "Q": "=G{r}*$F{r}",
    "R": "=H{r}*$F{r}",
    "S": "=L{r}*$F{r}",
    "T": "=M{r}*$F{r}",
}


def compiled_combo_rows(new_combos: list[tuple], first_new_row: int,
                        data_end: int, month_tag: str) -> dict[str, object]:
    """set_cells payload appending (company, partner, rate) rows that July's
    data contains but the hand-maintained Compiled list doesn't. The pivot
    caches read C4:T3723 / C4:T8724, so appended rows are inside range."""
    cells: dict[str, object] = {}
    r = first_new_row
    for company, partner, rate in new_combos:
        cells[f"B{r}"] = f"Auto {month_tag}"
        cells[f"C{r}"] = company
        cells[f"E{r}"] = partner
        cells[f"F{r}"] = rate
        for col, tpl in _CD_ROW_TPL.items():
            cells[f"{col}{r}"] = tpl.replace("{end}", str(data_end)) \
                                    .replace("{r}", str(r))
        r += 1
    return cells


# ── server-side rate engine (combo detection only) ────────────────────
# Mirrors the workbook's commission-rate formula. Used ONLY to decide which
# (company, partner, rate) rows to append to Compiled Data — the numbers on
# every sheet still come from the workbook's own formulas at recalc time,
# so a drift here can add a harmless zero row or miss a combo (reported in
# the job status for review), never change a computed amount.

def rate_for_row(partner, company, item_id, sku_rates: dict,
                 kevin_rates: dict, doc_no=None) -> float | None:
    p = (str(partner).strip() if partner is not None else "None")
    comp = str(company or "")
    if p in ("8 Nolita", "None"):
        return 0.0
    if p == "Kevin Hanks":
        kr = kevin_rates.get(str(doc_no).strip() if doc_no else "")
        return kr if isinstance(kr, (int, float)) else None
    sr = sku_rates.get(str(item_id).strip() if item_id is not None else "")
    if isinstance(sr, (int, float)):
        return sr
    if "Bomgaar" in comp:
        return None            # contracted-rate branch — skip, don't guess
    return 0.1                  # workbook default


def distinct_combos(prior_rows, cur_rows, sales_rows, sku_rates,
                    kevin_rates) -> tuple[set, set]:
    """-> (set of (company, partner, rate), set of (company, partner) whose
    rate could not be determined server-side)."""
    combos, undetermined = set(), set()
    for rows in (prior_rows, cur_rows, sales_rows):
        for row in rows:
            company = row[23] if len(row) > 23 else None     # X company (AR)
            if company is None and len(row) > 24:
                company = row[24]                            # sales Y
            partner = row[20]                                # U primary partner
            if not company or partner in (None, ""):
                continue
            comp = str(company).strip()
            # source dirt: numeric fragments ('0.02') land in the company
            # column on malformed rows — never seed Compiled rows off those
            if not re.search(r"[A-Za-z]{2}", comp):
                continue
            if str(partner).strip() in ("None", "8 Nolita"):
                continue
            r = rate_for_row(partner, company, row[8], sku_rates,
                             kevin_rates, doc_no=row[7])
            if r is None:
                undetermined.add((comp, str(partner)))
            elif r > 0:
                combos.add((comp, str(partner), round(float(r), 6)))
    return combos, undetermined


# ── per-rep statement files ────────────────────────────────────────────

_STMT_HEADERS = ["Client", "Commission %", "Prior Month AR", "New Sales",
                 "Collections", "Partial Payments", "Total Collections",
                 "Commission Earned"]


def _mini_xlsx(sheet_name: str, rows: list[list]) -> bytes:
    """A minimal single-sheet .xlsx built from plain values."""
    from surgeon import row_xml, col_letter   # reuse the proven serializers
    import io
    body = "".join(row_xml(i + 1, r) for i, r in enumerate(rows))
    ncols = max((len(r) for r in rows), default=1)
    dim = f"A1:{col_letter(ncols)}{max(len(rows), 1)}"
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main">'
             f'<dimension ref="{dim}"/><sheetData>{body}</sheetData></worksheet>')
    esc_name = (sheet_name.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))[:31]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   f'<sheets><sheet name="{esc_name}" sheetId="1" r:id="rId1"/></sheets>'
                   '<calcPr calcId="1"/></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                   '</Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()


def build_statements(compiled: list[dict], payment: dict[str, dict],
                     period_label: str) -> dict[str, bytes]:
    """{filename: xlsx bytes} — one workbook per partner.

    compiled rows: {company, partner, rate, prior, newSales, collections,
    partial, totalColl, earned} (post-recalc cached values). payment:
    {partner: {earned, fee, adj, net}}. Kevin Hanks is presented by client
    only (rates aggregated) per the SOP exception."""
    by_partner: dict[str, list[dict]] = {}
    for row in compiled:
        if not row.get("partner"):
            continue
        by_partner.setdefault(str(row["partner"]), []).append(row)

    out: dict[str, bytes] = {}
    for partner, prows in sorted(by_partner.items()):
        if partner in ("None", "Dayna Stambeck & Luxe Brands"):
            continue
        if all(abs(r.get("earned") or 0) < 0.005 and
               abs(r.get("totalColl") or 0) < 0.005 for r in prows):
            continue                       # nothing to report this month
        kevin = partner == "Kevin Hanks"
        if kevin:                          # by client only
            agg: dict[str, dict] = {}
            for r in prows:
                a = agg.get(str(r["company"]))
                if a is None:
                    agg[str(r["company"])] = dict(r, rate=None)
                else:
                    for k in ("prior", "newSales", "collections", "partial",
                              "totalColl", "earned"):
                        a[k] = (a.get(k) or 0) + (r.get(k) or 0)
            prows = list(agg.values())
        rows: list[list] = [
            [f"Commission Statement — {partner}"],
            [f"Period: {period_label}"],
            [],
            list(_STMT_HEADERS),
        ]
        tot = dict.fromkeys(("prior", "newSales", "collections", "partial",
                             "totalColl", "earned"), 0.0)
        for r in sorted(prows, key=lambda x: str(x["company"])):
            if all(abs(r.get(k) or 0) < 0.005 for k in tot):
                continue
            rows.append([
                r["company"],
                None if kevin or r.get("rate") in (None, "") else r["rate"],
                round(r.get("prior") or 0, 2), round(r.get("newSales") or 0, 2),
                round(r.get("collections") or 0, 2), round(r.get("partial") or 0, 2),
                round(r.get("totalColl") or 0, 2), round(r.get("earned") or 0, 2),
            ])
            for k in tot:
                tot[k] += r.get(k) or 0
        rows.append(["Total", None] + [round(tot[k], 2) for k in
                     ("prior", "newSales", "collections", "partial",
                      "totalColl", "earned")])
        pay = payment.get(partner)
        if pay:
            rows += [[],
                     ["Commission Earned (per Summary)", round(pay.get("earned") or 0, 2)],
                     ["Technology Fee", round(pay.get("fee") or 0, 2)],
                     ["Other Adjustments", round(pay.get("adj") or 0, 2)],
                     ["Net Payment", round(pay.get("net") or 0, 2)]]
        safe = re.sub(r'[\\/:*?"<>|]', "_", partner)
        out[f"{period_label}_Commission Statement_{safe}.xlsx"] = \
            _mini_xlsx("Statement", rows)
    return out
