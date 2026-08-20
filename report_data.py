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


# full-width row layout: A(0) DateBreak | B(1) Period | C..Z(2..25) data |
# AA(26) Difference | AB(27) Deposit Month | AC(28) Client Age | AD(29)
# Deposit Month # | AE(30) Rate | AF(31) Unearned | AG(32) Earned | AH(33)
# unused | AI(34) key | AJ(35) prior open bal | AK(36) gross original |
# AL(37) gross adjusted | AM(38) key2 | AN(39)/AO(40) left empty
DATA_WIDTH = 39


def build_data_rows(prior_rows: list[list], cur_rows: list[list],
                    sales_rows: list[list], asof_serial: int,
                    prior_tab: str, cur_tab: str, sales_tab: str,
                    earned_label: str) -> dict:
    """-> {'pasteRows': [...full-width A..AM rows, formulas inline...],
          'lastRow', 'blocks', 'headerCells'}.

    Formulas ride INSIDE the paste rows ('=' strings) — the first cut kept
    them in a separate ~330k-entry {ref: formula} dict, and that dict plus
    the paste rows plus phase 2's still-resident ops OOM'd the container
    (0820 evening). CONSUMES the source lists (cleared block by block) so
    peak memory is roughly the output alone."""
    paste: list[list] = []
    rn = DATA_FIRST_ROW

    def is_num(v):
        return isinstance(v, (int, float))

    def base(label, vals24):
        r = [None] * DATA_WIDTH
        r[1] = label
        r[2:2 + 24] = list(vals24[:24]) + [None] * (24 - len(vals24[:24]))
        return r

    # ---- block 1: Prior Month ------------------------------------------
    blk1_first = rn
    n1 = len(prior_rows)
    for i in range(n1):
        row = prior_rows[i]
        prior_rows[i] = None                    # consume as we go
        src = 7 + i                             # tab data starts at row 7
        r = base("Prior Month", row)
        t = row[19]                             # verified Date Closed serial
        if is_num(t) and t <= asof_serial:
            r[27] = t                           # SOP: future months removed
            r[29] = _serial_month(t)
        else:
            r[29] = 0
        fs = row[3]
        if is_num(fs):
            r[28] = round((asof_serial - fs) / 365, 2)
        r[30] = f"='{prior_tab}'!AD{src}"       # rate
        r[31] = f"='{prior_tab}'!AE{src}"       # unearned
        r[32] = f"='{prior_tab}'!AF{src}"       # earned
        r[34] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        r[38] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        paste.append(r)
        rn += 1
    prior_rows.clear()
    blk1_last = rn - 1
    for _ in range(SPACER):                     # paste is consecutive from
        paste.append([None] * DATA_WIDTH)       # the anchor — spacers must
        rn += 1                                 # be real rows

    # ---- block 2: Current Month ----------------------------------------
    blk2_first = rn
    aj_rng = f"$AI${blk1_first}:$AI${blk1_last}"
    aj_n = f"$N${blk1_first}:$N${blk1_last}"
    n2 = len(cur_rows)
    for i in range(n2):
        row = cur_rows[i]
        cur_rows[i] = None
        src = 7 + i
        r = base("Current Month", row)
        r[26] = f"=X{rn}-N{rn}"                 # SOP: formula, not value
        r[30] = f"='{cur_tab}'!AA{src}"         # rate (June layout)
        r[31] = f"='{cur_tab}'!AB{src}"         # unearned
        r[34] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        r[35] = (f'=IF($B{rn}="Current Month",_xlfn.XLOOKUP('
                 f'AI{rn},{aj_rng},{aj_n},"Error",0),0)')
        gross = row[21]                         # V Amount (Gross)
        if gross is not None:
            r[36] = gross                       # Original, as VALUE
        r[37] = (f"=IF(AJ{rn}<0,IF(AJ{rn}>AK{rn},AJ{rn},"
                 f"IF(AJ{rn}<AK{rn},AJ{rn},AK{rn})),"
                 f"IF(AJ{rn}<AK{rn},AJ{rn},AK{rn}))")
        r[23] = f"=AL{rn}"      # X: the paste-adjusted-over-gross step
        r[38] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        paste.append(r)
        rn += 1
    cur_rows.clear()
    blk2_last = rn - 1
    for _ in range(SPACER):
        paste.append([None] * DATA_WIDTH)
        rn += 1

    # ---- block 3: New Sales --------------------------------------------
    blk3_first = rn
    n3 = len(sales_rows)
    for i in range(n3):
        row = sales_rows[i]
        sales_rows[i] = None
        src = 7 + i
        r25 = list(row[:25]) + [None] * (25 - len(row[:25]))
        r = base("New Sales", r25[:22])
        r[25] = r25[24]                         # Z <- sales Y (company)
        r[27] = f"='{sales_tab}'!AB{src}"       # deposit date (gated)
        r[28] = f"='{sales_tab}'!X{src}"        # client age
        r[29] = f"='{sales_tab}'!AC{src}"       # deposit month number
        r[30] = f"='{sales_tab}'!AD{src}"       # rate
        r[31] = f"='{sales_tab}'!AE{src}"       # unearned
        r[32] = f"='{sales_tab}'!AF{src}"       # earned
        r[34] = f"=_xlfn.CONCAT(Z{rn},J{rn},K{rn})"
        r[38] = f'=_xlfn.CONCAT(J{rn}," - ",K{rn})'
        paste.append(r)
        rn += 1
    sales_rows.clear()
    blk3_last = rn - 1

    return {
        "pasteRows": paste, "lastRow": blk3_last,
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


# ── combo detection (drives Compiled row appends) ─────────────────────
# Uses the same shadow_rate engine that penny-matched Mike's June book, so
# a combo exists exactly when the workbook's own formulas will produce a
# nonzero rate for it.

def distinct_combos(prior_rows, cur_rows, sales_rows, sku_rates,
                    licensed_ids, consts) -> tuple[set, set]:
    """-> (set of (company, partner, rate), set of (company, partner) whose
    rate resolved to an excluded/blank value)."""
    combos, undetermined = set(), set()
    for rows, comp_idx in ((prior_rows, 23), (cur_rows, 23), (sales_rows, 24)):
        for row in rows:
            company = row[comp_idx] if len(row) > comp_idx else None
            partner = row[20]
            if not company or partner in (None, ""):
                continue
            comp = str(company).strip()
            # source dirt: numeric fragments ('0.02') land in the company
            # column on malformed rows — never seed Compiled rows off those
            if not re.search(r"[A-Za-z]{2}", comp):
                continue
            if str(partner).strip() in ("None", "8 Nolita"):
                continue
            r = shadow_rate(partner, company, row[8], row[7], row[5],
                            sku_rates, licensed_ids, consts)
            if isinstance(r, (int, float)) and r > 0:
                combos.add((comp, str(partner), round(float(r), 6)))
            elif r == " ":
                undetermined.add((comp, str(partner)))
    return combos, undetermined


# ── shadow calc: the workbook's commission arithmetic, ported 1:1 ─────
# Statements need Compiled-level numbers, which Excel only computes at
# recalc time. Rather than gate statements on a desktop round-trip, the
# service computes them with the SAME rules (formulas ported line for
# line from FORMULA_TEMPLATES / the June book) and the port is validated
# by replaying June's data and matching Mike's own cached Summary/Payment
# values. The workbook keeps its real formulas — recalc remains the
# authority; this mirror exists so the flow can run end to end.

def _contracted_stack(company, partner, doc_no, f_serial, licensed,
                      ai1_serial) -> float | str:
    """AR!Y / sales!AA — SEARCH() is case-insensitive."""
    co = str(company or "").lower()
    year = (EXCEL_EPOCH + dt.timedelta(days=int(f_serial))).year \
        if isinstance(f_serial, (int, float)) else None
    doc = None
    try:
        doc = float(doc_no)
    except (TypeError, ValueError):
        pass
    if "cavenders" in co and doc is not None and doc > 1121738 \
            and year is not None and year < 2026:
        return 0.09
    if "cavenders" in co:
        return 0.0275
    if "scheels" in co and doc is not None and doc > 1122227 \
            and year is not None and year < 2026:
        return 0.07
    if any(s in co for s in ("atwoods", "scheels", "the glik company",
                             "quiet storm", "glik")):
        return 0.05
    if "bomgaar" in co and str(partner) == "Kelly Kennedy" \
            and isinstance(f_serial, (int, float)) and f_serial > ai1_serial:
        return 0.045 if licensed else 0.06
    return " "


def shadow_rate(partner, company, item_id, doc_no, f_serial,
                sku_rates: dict, licensed_ids: set,
                consts: dict) -> float | str:
    """AR!AA / sales!AD. Returns a float, '0' (excluded partners), or ' '
    (excluded rows). IFERROR fallback -> 0.1."""
    u = str(partner).strip() if partner is not None else "None"
    if u in ("8 Nolita", "None"):
        return "0"
    licensed = str(item_id).strip() in licensed_ids
    if u == "Kevin Hanks":
        return _contracted_stack(company, u, doc_no, f_serial, licensed,
                                 consts["ai1"])
    if "bomgaars" in str(company or "").lower():
        return _contracted_stack(company, u, doc_no, f_serial, licensed,
                                 consts["ai1"])
    if u == consts["dayna"] and isinstance(f_serial, (int, float)) \
            and f_serial > consts["ad1"]:
        return " "
    sr = sku_rates.get(str(item_id).strip() if item_id is not None else "")
    if isinstance(sr, (int, float)):
        return sr
    return 0.1                                   # VLOOKUP miss -> IFERROR


def _sales_deposit(u, g_type, f_serial, t_serial, asof_serial,
                   consts: dict):
    """sales!AB (IFS): the deposit date, or None. AJ4 (45762) caps the
    Dayna exception; Cash Sales deposit on their own date."""
    if u == consts["dayna"] and isinstance(t_serial, (int, float)) \
            and t_serial > consts["aj4"]:
        return None
    if str(g_type or "") == "Cash Sale":
        return f_serial if isinstance(f_serial, (int, float)) else None
    if not isinstance(t_serial, (int, float)):
        return None
    return t_serial if t_serial <= asof_serial else None


def shadow_compiled(prior_rows, cur_rows, sales_rows, asof_serial: int,
                    sku_rates: dict, licensed_ids: set,
                    consts: dict) -> list[dict]:
    """Compiled Data rows 5+ recomputed: per (company, partner, rate) —
    G prior, H new sales, I/J collections (Compiled's negative-sum sign),
    K partial payments, L total collections, S earned = L*rate."""
    def key_of(row):
        return f"{row[23]}{row[7]}{row[8]}"          # CONCAT(Z, J, K)

    prior_bal = {}
    for row in prior_rows:
        prior_bal.setdefault(key_of(row), row[11])   # XLOOKUP = first match

    agg: dict[tuple, dict] = {}

    def bucket(company, partner, rate):
        k = (str(company), str(partner), round(float(rate), 6))
        return agg.setdefault(k, dict.fromkeys(
            ("G", "H", "I", "J", "K"), 0.0))

    for row in prior_rows:                            # 'Prior Month' block
        company, partner = row[23], row[20]
        rate = shadow_rate(partner, company, row[8], row[7], row[5],
                           sku_rates, licensed_ids, consts)
        if not isinstance(rate, (int, float)) or rate <= 0:
            continue
        n = row[11] if isinstance(row[11], (int, float)) else 0.0
        b = bucket(company, partner, rate)
        b["G"] += n
        t = row[19]                                   # verified Date Closed
        if isinstance(t, (int, float)) and t <= asof_serial:
            b["I"] -= n

    for row in cur_rows:                              # 'Current Month' block
        company, partner = row[23], row[20]
        rate = shadow_rate(partner, company, row[8], row[7], row[5],
                           sku_rates, licensed_ids, consts)
        if not isinstance(rate, (int, float)) or rate <= 0:
            continue
        n = row[11] if isinstance(row[11], (int, float)) else 0.0
        ak = row[21] if isinstance(row[21], (int, float)) else 0.0
        aj = prior_bal.get(f"{company}{row[7]}{row[8]}")
        if not isinstance(aj, (int, float)):
            al = ak                                   # XLOOKUP 'Error' path
        elif aj < 0:
            al = aj
        else:
            al = min(aj, ak)
        b = bucket(company, partner, rate)
        b["K"] -= (al - n)                            # K = -sum(Difference)

    for row in sales_rows:                            # 'New Sales' block
        company, partner = row[24], row[20]
        rate = shadow_rate(partner, company, row[8], row[7], row[5],
                           sku_rates, licensed_ids, consts)
        if not isinstance(rate, (int, float)) or rate <= 0:
            continue
        n = row[11] if isinstance(row[11], (int, float)) else 0.0
        b = bucket(company, partner, rate)
        b["H"] += n
        dep = _sales_deposit(str(partner).strip() if partner else "",
                             row[6], row[5], row[19], asof_serial, consts)
        if isinstance(dep, (int, float)) and dep > 0:
            b["J"] -= n

    out = []
    for (company, partner, rate), b in sorted(agg.items()):
        total = b["I"] + b["J"] + b["K"]
        out.append({"company": company, "partner": partner, "rate": rate,
                    "prior": b["G"], "newSales": b["H"],
                    "collections": b["I"] + b["J"], "partial": b["K"],
                    "totalColl": total, "earned": total * rate})
    return out


def shadow_payment(compiled: list[dict], fee_table: dict) -> dict:
    """Payment tab: per-partner earned (Summary XLOOKUP) + tech fee ->
    net = min(0, earned + fee). Other Adjustments stay manual (0)."""
    earned: dict[str, float] = {}
    for r in compiled:
        earned[r["partner"]] = earned.get(r["partner"], 0.0) + r["earned"]
    out = {}
    for partner, e in earned.items():
        fee = fee_table.get(partner, 0.0)
        net = e + fee
        out[partner] = {"earned": e, "fee": fee, "adj": 0.0,
                        "net": net if net <= 0 else 0.0}
    return out


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
