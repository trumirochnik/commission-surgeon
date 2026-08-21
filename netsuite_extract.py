"""
netsuite_extract.py — commission extract for Tru Fragrance.

Pulls the two datasets the commission workbook needs, via the NetSuite MCP.
Runs standalone for reconciliation, or is imported by service.py.

    python netsuite_extract.py --asof 2026-07-31 --from 2026-07-01 --to 2026-07-31

RECONCILIATION TARGETS (measured against acct 430465 for July 2026):
    AR    : 16,362 rows   open balance sum $2,680,818.23
    Sales : ~5,165 transactions -> ~13,500 lines

════════════════════════════════════════════════════════════════════════
RULES THAT ARE NOT NEGOTIABLE — every one was established by measurement.
Changing any of them reintroduces a failure that already cost days.
════════════════════════════════════════════════════════════════════════

1. Open-at-as-of is `trandate <= ASOF AND (closedate IS NULL OR closedate > ASOF)`.
   16,362 rows in seconds. The obvious alternative — deriving open balance from
   an unbounded NextTransactionLineLink CTE — DOES NOT RETURN. Measured >4min.

2. The line query joins ONLY transaction + transactionline + item.
   Joining `customer` or `transactionShippingAddress` collapses the query plan:
   60 transaction ids timed out at 60s. Customer attributes come from a separate
   dimension query keyed by id. This is the single most important rule here.

3. The customer dimension is fetched BY ID, for only the ids seen on the lines.
   `SELECT ... FROM customer` unfiltered is 50,712 rows x 4 BUILTIN.DF calls
   and never finishes.

4. The MCP enforces a 1000-row page limit and REQUIRES offsets to be exact
   multiples of 1000. `next_offset` is row-count based, so it must NOT be fed
   back as an offset (offset 3 is rejected). Advance by PAGE and treat a short
   page as end-of-data.

5. CustInvc lines return NEGATIVE quantity and amount on this account
   (verified: tran 29506822 -> qty -744, amt -4166.40). sign_flip corrects it.
   RECONCILE against a hand-built month before trusting the numbers.

6. Commission % is left EMPTY. The workbook's "Commission Rate by SKUs" lookup
   owns that column. Do not compute it here.

COLUMN SHAPES (read off the 06.2026 workbook 2026-08-17):
    AR tab            A:X  = 24 cols, header row 6, data row 7,  formulas Y:AH
    New Sales report  A:Y  = 25 cols, header row 6, data row 7,  formulas Z:AH
    Sales report Raw  A:Y  = 25 cols, header row 3, data row 4,  formulas Z:AA
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
from typing import Any, Iterable

import requests

PAGE = 1000                  # rule 4: fixed, and offsets must be multiples of it
ID_BATCH = 250               # transactions per line-detail call
CUST_BATCH = 400             # customer ids per dimension call
EXCEL_EPOCH = dt.date(1899, 12, 30)

# Column P ("Partner: Partner Category") comes from the partner record's category.
# VERIFIED 2026-08-17 once Lists > Partners (View) was granted: 74 partners, and the
# value is NOT uniform — "Primary Rep" mostly, but also "Rep Manager" (Gary Pollack),
# "Royalty" (BATA), "Adjustment Partner", and blanks. Hardcoding it would be wrong.
PARTNER_ROLE_FALLBACK = None

# Column M renders as "EMP101 Mark Saviski" — code AND name. BUILTIN.DF(t.employee)
# gives only the code, and the `employee` table is permission-blocked for this
# credential (HTTP 400 "Record not found" = missing Lists > Employee Record View).
# There are only 13 distinct codes in a month, so this is a lookup table.
# Unmapped codes fall through as the bare code and are listed in diagnostics —
# read the missing names off column M of the workbook and add them here.
EMPLOYEE_NAMES = {
    "EMP101": "Mark Saviski",
    "EMP140": "Not Applicable",
    "EMP204": "Kevin Hanks",
    # "EMP109": "", "EMP124": "", "EMP161": "", "EMP165": "", "EMP179": "",
    # "EMP208": "", "EMP245": "", "EMP254": "", "EMP262": "", "EMP374": "",
}


def employee_label(code: Any) -> Any:
    """'EMP101' -> 'EMP101 Mark Saviski'; unknown codes pass through unchanged."""
    if not code:
        return code
    name = EMPLOYEE_NAMES.get(str(code).strip())
    return f"{code} {name}" if name else code

EXCLUDED_PARTNERS = {
    "Pfeiffer of America", "RMCC", "RMCC/Canada", "RMCC/Cavenders",
    "RMCC/Gander Mountain", "RMCC/Midstates", "RMCC/Wheatbelt", "Temp",
    "Erin Clucas", "BrendaCollins", "Brenda Collins",
    "Atlanta Apparel Expo", "BATA", "Tom Tompkins",
}
EXCLUDED_COMPANIES = {"Amazon", "Amazon - Undone", "Shop.Tru", "Undone Shopify"}
STORE_TYPES = (
    "Department Store", "Gift Shop", "Mens/Womens", "Menswear",
    "Western", "Womenswear", "Pro Shop", "Boutique",
)
ITEM_TYPES = ("Assembly", "Discount", "InvtPart")
TXN_TYPES = ("CustInvc", "CashSale", "CustCred")
AR_HISTORY_START = "2022-10-01"


# ─────────────────────────── MCP client ───────────────────────────

class McpError(RuntimeError):
    pass


class Mcp:
    """Thin SuiteQL client. Enforces rule 4 so callers cannot get it wrong."""

    def __init__(self, url: str, secret: str, header: str = "x-mcp-secret",
                 timeout: int = 300):
        self.url = url
        self.timeout = timeout
        self.headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
            header: secret,
        }
        self._id = 0

    def _call(self, query: str, offset: int) -> dict:
        self._id += 1
        body = {
            "jsonrpc": "2.0", "id": self._id, "method": "tools/call",
            "params": {"name": "netsuite_run_suiteql", "arguments": {
                "query": query, "limit": PAGE, "offset": offset, "format": "json"}},
        }
        r = requests.post(self.url, headers=self.headers, json=body,
                          timeout=self.timeout)
        if r.status_code == 401:
            raise McpError("MCP 401 — check the shared secret / header name.")
        if r.status_code >= 400:
            raise McpError(f"MCP HTTP {r.status_code}: {r.text[:400]}")

        payload = self._unwrap(r.text)
        if payload is None:
            raise McpError(f"Unparseable MCP response: {r.text[:400]}")
        return payload

    @staticmethod
    def _unwrap(text: str) -> dict | None:
        """The MCP answers as JSON or as SSE `data:` lines. Handle both."""
        env = None
        stripped = text.strip()
        if stripped.startswith("{"):
            try:
                env = json.loads(stripped)
            except ValueError:
                env = None
        if env is None:
            for line in re.findall(r"^data:\s?(.*)$", text, re.M):
                try:
                    env = json.loads(line)
                    break
                except ValueError:
                    continue
        if env is None:
            return None
        if env.get("error"):
            raise McpError(f"MCP error: {json.dumps(env['error'])[:400]}")
        result = env.get("result") or {}
        if result.get("isError"):
            raise McpError(f"SuiteQL error: {json.dumps(result)[:600]}")
        if "structuredContent" in result:
            return result["structuredContent"]
        blocks = [b.get("text", "") for b in result.get("content", [])
                  if b.get("type") == "text"]
        if blocks:
            try:
                return json.loads("".join(blocks))
            except ValueError:
                raise McpError(f"Tool result was not JSON: {''.join(blocks)[:400]}")
        return result

    def rows(self, query: str, label: str = "query") -> list[dict]:
        """Page through a query. Offsets advance by PAGE — never by next_offset."""
        out: list[dict] = []
        offset = 0
        for page in range(400):
            payload = self._call(query, offset)
            batch = payload.get("rows")
            if batch is None:
                raise McpError(f"{label}: unexpected shape {json.dumps(payload)[:300]}")
            out.extend(batch)
            if len(batch) < PAGE:          # rule 4: short page == done
                return out
            offset += PAGE
        raise McpError(f"{label}: page cap hit at {len(out)} rows")


# ─────────────────────────── helpers ───────────────────────────

def to_date(v: Any) -> dt.date | None:
    """Parse NetSuite's M/D/YYYY or ISO date strings."""
    if v in (None, ""):
        return None
    s = str(v)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def serial(v: Any) -> Any:
    """Excel serial date. NetSuite hands back M/D/YYYY or ISO."""
    d = to_date(v)
    if d is None:
        return None if v in (None, "") else v
    return (d - EXCEL_EPOCH).days


_MON3 = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def month_label(v: Any) -> Any:
    """'6/15/2026' -> 'Jun 2026' — the sales tabs' W column format, matched
    to existing rows ('Oct 2022', 'Jun 2026'). Manual month names, not
    strftime, so a non-English server locale can't change the output."""
    d = to_date(v)
    return f"{_MON3[d.month - 1]} {d.year}" if d else None


def num(v: Any) -> Any:
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def g(row: dict, n: int) -> Any:
    k = f"c{n:02d}"
    return row.get(k, row.get(k.upper()))


def chunks(seq: list, size: int) -> Iterable[list]:
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def sql_list(vals: Iterable[str]) -> str:
    return ", ".join("'" + str(v).replace("'", "''") + "'" for v in vals)


# ─────────────────────────── queries ───────────────────────────

def q_ar(asof: str) -> str:
    """AR aging detail, open as of `asof`.

    Rule 1 supplies the open-at-as-of test. The as-of BALANCE is the header's
    current unpaid amount plus anything paid AFTER as-of (a small bounded CTE —
    ~3,328 docs — not the unbounded one that hangs), prorated across lines.
    """
    return f"""
WITH postpay AS (
  SELECT ntll.previousdoc AS tid, SUM(NVL(ntll.foreignamount,0)) AS post_amt
  FROM NextTransactionLineLink ntll
  JOIN transaction pt ON pt.id = ntll.nextdoc
  WHERE pt.trandate > TO_DATE('{asof}','YYYY-MM-DD')
  GROUP BY ntll.previousdoc
)
SELECT
  BUILTIN.DF(t.entity) AS c01, t.trandate AS c02, BUILTIN.DF(t.type) AS c03,
  t.tranid AS c04, i.itemid AS c05, tl.quantity AS c06,
  ROUND(tl.foreignamount * (NVL(hl.foreignamountunpaid,0) + NVL(pp.post_amt,0))
        / NULLIF(t.foreigntotal,0), 2) AS c07,
  BUILTIN.DF(t.employee) AS c08, BUILTIN.DF(t.partner) AS c09,
  BUILTIN.DF(t.status) AS c10, t.duedate AS c11, t.closedate AS c12,
  tl.foreignamount AS c13, t.id AS c14, t.entity AS c15
FROM transaction t
JOIN transactionline tl ON tl.transaction = t.id AND tl.mainline = 'F'
JOIN item i ON i.id = tl.item
LEFT JOIN transactionline hl ON hl.transaction = t.id AND hl.mainline = 'T'
LEFT JOIN postpay pp ON pp.tid = t.id
WHERE t.type IN ({sql_list(TXN_TYPES)})
  AND t.trandate >= TO_DATE('{AR_HISTORY_START}','YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{asof}','YYYY-MM-DD')
  AND (t.closedate IS NULL OR t.closedate > TO_DATE('{asof}','YYYY-MM-DD'))
  AND tl.foreignamount IS NOT NULL
  AND i.itemtype IN ({sql_list(ITEM_TYPES)})
ORDER BY t.id, tl.id
""".strip()


def q_sales_ids(frm: str, to: str) -> str:
    """Header-only id sweep. Cheap: no joins, ~5,165 rows for July 2026.
    `t.partner IS NOT NULL` on the raw column — NOT BUILTIN.DF(...) IS NOT NULL,
    which is a per-row function call and cannot use an index."""
    return f"""
SELECT t.id AS c01
FROM transaction t
WHERE t.type IN ({sql_list(TXN_TYPES)})
  AND t.trandate BETWEEN TO_DATE('{frm}','YYYY-MM-DD') AND TO_DATE('{to}','YYYY-MM-DD')
  AND t.partner IS NOT NULL
ORDER BY t.id
""".strip()


def q_sales_lines(ids: list[str]) -> str:
    """Rule 2: transaction + transactionline + item ONLY.
    (tl.rate was dropped 2026-08-18 — it was only ever written into the
    sales tabs' Y column, which is actually 'Company name' in the layout.)"""
    return f"""
SELECT
  BUILTIN.DF(t.entity) AS c01, t.trandate AS c02, BUILTIN.DF(t.type) AS c03,
  t.tranid AS c04, i.itemid AS c05, tl.quantity AS c06, tl.foreignamount AS c07,
  BUILTIN.DF(t.employee) AS c08, BUILTIN.DF(t.partner) AS c09,
  BUILTIN.DF(t.status) AS c10, t.duedate AS c11, t.closedate AS c12,
  t.id AS c14, t.entity AS c15
FROM transaction t
JOIN transactionline tl ON tl.transaction = t.id AND tl.mainline = 'F'
JOIN item i ON i.id = tl.item
WHERE t.id IN ({','.join(ids)})
  AND tl.foreignamount IS NOT NULL
  AND i.itemtype IN ({sql_list(ITEM_TYPES)})
ORDER BY t.id, tl.id
""".strip()


def q_customers(ids: list[str]) -> str:
    """Rule 3: by id only.

    VERIFIED 2026-08-17: `custentity_shipping_state` and `partnercategory` DO NOT
    EXIST on customer (HTTP 500). Shipping state comes from fetch_states() instead.
    The partner ROLE (col P, "Primary Rep") is not reachable — the `partner` and
    `entity` tables are permission-blocked for this role — so PARTNER_ROLE is a
    constant. See those notes below.
    """
    return f"""
SELECT cu.id AS c01,
  BUILTIN.DF(cu.custentitycustomer_type) AS c02,
  BUILTIN.DF(cu.category) AS c03,
  cu.firstsaledate AS c04,
  BUILTIN.DF(cu.custentitystore_type) AS c05,
  cu.companyname AS c06,
  BUILTIN.DF(cu.partner) AS c07,
  BUILTIN.DF(cu.custentity12) AS c08
FROM customer cu
WHERE cu.id IN ({','.join(ids)})
ORDER BY cu.id
""".strip()


def q_items(ids: list[str]) -> str:
    """Real product descriptions for column J.
    VERIFIED 2026-08-17: `i.description` works (`i.salesdescription` is HTTP 500).
    e.g. itemid 96993 -> "TJX - Women's Elements Overspray - 100 mL EDP - Ruby".

    Keyed on i.itemid, NOT i.id: the line queries return c05 = i.itemid (the SKU
    code). SKUs are not all numeric — there are codes like DISC00 and OLD* — so they
    MUST be quoted, or SuiteQL parses them as identifiers and rejects the query."""
    return f"""
SELECT i.itemid AS c01, i.description AS c02
FROM item i
WHERE i.itemid IN ({sql_list(ids)})
""".strip()


# ─────────────────────────── dimensions ───────────────────────────

def fetch_customers(mcp: Mcp, ids: list[str]) -> dict[str, dict]:
    """Rule 3: by id, for only the ids seen on the lines.

    commission_pct: customer field custentity12, a SELECT field whose option
    LABELS are the whole-number percents {5,7,10,14,15,20} — VERIFIED
    2026-08-18 against five customers whose saved-search "Commission Pct"
    values were known (Johnson's 10, American Man 20, Cavenders#34 5,
    Cattleman 14, Walker 7). BUILTIN.DF gives the label; the raw column is
    the option's internal id (1..6) and must NOT be used directly.

    Kept as the label TEXT ("10", not 0.10): the workbook has always stored
    Q as text whole numbers — the AGA saved-search CSV export shows literal
    "5"/"10"/"20"/"7" — and downstream consumers expect that. (An earlier
    revision converted to fractions; that conversion was OURS, not
    NetSuite's, and it was wrong to add.) Nothing in the Y:AH formula block
    references Q, so text is safe."""
    def pct(label):
        s = str(label).strip() if label not in (None, "") else ""
        return s or None

    out: dict[str, dict] = {}
    for batch in chunks(ids, CUST_BATCH):
        for r in mcp.rows(q_customers(batch), "customer dim"):
            out[str(g(r, 1))] = {
                "type": g(r, 2), "category": g(r, 3),
                "first_sale": serial(g(r, 4)), "store_type": g(r, 5),
                "company": g(r, 6), "partner": g(r, 7),
                "commission_pct": pct(g(r, 8)),
            }
    return out


def q_partners() -> str:
    """Partner name -> category, for column P. 74 rows, one page.
    Keyed on entityid because the line query returns BUILTIN.DF(t.partner) (a name),
    not the internal id. Requires Lists > Partners (View)."""
    return "SELECT p.entityid AS c01, BUILTIN.DF(p.category) AS c02 FROM partner p"


def fetch_partners(mcp: Mcp) -> dict[str, str]:
    out: dict[str, str] = {}
    for r in mcp.rows(q_partners(), "partner dim"):
        name, cat = g(r, 1), g(r, 2)
        if name and cat:
            out[str(name).strip()] = cat
    return out


def q_states(ids: list[str]) -> str:
    """Shipping state (col N) keyed by TRANSACTION id.

    VERIFIED 2026-08-17: transactionShippingAddress.state works fine on a small
    keyed id set (returns IL/TN/NM). It is only fatal when joined into the full
    line query — see rule 2. The customer's own address tables
    (CustomerAddressbook*) return empty for this role, and `defaultshippingaddress`
    is just an id with nothing readable behind it, so this is the only route.
    """
    return f"""
SELECT t.id AS c01, tsa.state AS c02
FROM transaction t
LEFT JOIN transactionShippingAddress tsa ON tsa.nkey = t.shippingaddress
WHERE t.id IN ({','.join(ids)})
""".strip()


def fetch_states(mcp: Mcp, ids: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for batch in chunks(ids, CUST_BATCH):
        for r in mcp.rows(q_states(batch), "shipping state"):
            st = g(r, 2)
            if st:
                out[str(g(r, 1))] = st
    return out


def q_accounts(ids: list[str]) -> str:
    """AR column W 'Account: Name (GL-style)' keyed by TRANSACTION id.

    VERIFIED 2026-08-18: the transaction HEADER carries the receivables
    account directly — BUILTIN.DF(t.account) = '11300 Accounts Receivable -
    Trade' for tran 29506822 — so no transactionaccountingline join is
    needed at all (putting that join in the line query is what collapsed
    the plan originally). Keyed-id batches only, same pattern as
    fetch_states(); the `account` table itself is permission-blocked for
    this credential, which is fine since DF does the naming."""
    return f"""
SELECT t.id AS c01, BUILTIN.DF(t.account) AS c02
FROM transaction t
WHERE t.id IN ({','.join(ids)})
""".strip()


def fetch_accounts(mcp: Mcp, ids: list[str]) -> dict[str, str]:
    """{txn id: '11300 - Accounts Receivable - Trade'}. BUILTIN.DF returns
    '11300 Accounts Receivable - Trade'; the workbook/saved-search format
    has a dash after the number (ground truth: the AGA saved-search CSV) —
    inserted here."""
    out: dict[str, str] = {}
    for batch in chunks(ids, CUST_BATCH):
        for r in mcp.rows(q_accounts(batch), "AR account dim"):
            name = g(r, 2)
            if name:
                out[str(g(r, 1))] = re.sub(r"^(\d+)\s+", r"\1 - ", str(name))
    return out


def fetch_new_items(mcp: Mcp, frm: str, to: str, log=print) -> tuple[list, str]:
    """SKUs whose FIRST item fulfillment falls inside the period — 'first
    fulfilled', not merely 'sold this month'. Two steps, both verified live
    2026-08-19: (1) DISTINCT SKUs shipped (ItemShip) in the period — 637 for
    July; (2) keyed batches, MIN(trandate) over ALL fulfillment history with
    a server-side HAVING so only period-first SKUs come back (July: 99416,
    99417). Falls back to first appearance in SALES history if the role
    loses fulfillment visibility; the basis string labels which ran."""
    def q_shipped(kind_types):
        return f"""
SELECT DISTINCT i.itemid AS c01
FROM transaction t
JOIN transactionline tl ON tl.transaction = t.id AND tl.mainline = 'F'
JOIN item i ON i.id = tl.item
WHERE t.type IN ({kind_types})
  AND t.trandate BETWEEN TO_DATE('{frm}','YYYY-MM-DD') AND TO_DATE('{to}','YYYY-MM-DD')
  AND i.itemtype IN ('Assembly','InvtPart')
ORDER BY i.itemid
""".strip()

    def q_first(kind_types, ids):
        return f"""
SELECT i.itemid AS c01, MIN(t.trandate) AS c02
FROM transaction t
JOIN transactionline tl ON tl.transaction = t.id AND tl.mainline = 'F'
JOIN item i ON i.id = tl.item
WHERE t.type IN ({kind_types})
  AND i.itemid IN ({sql_list(ids)})
GROUP BY i.itemid
HAVING MIN(t.trandate) >= TO_DATE('{frm}','YYYY-MM-DD')
ORDER BY i.itemid
""".strip()

    for kinds, basis in (("'ItemShip'", "item_fulfillment"),
                         (sql_list(TXN_TYPES), "first_sale_fallback")):
        try:
            shipped = [str(g(r, 1)) for r in
                       mcp.rows(q_shipped(kinds), f"new items sweep ({basis})")]
            log(f"[newitems] {len(shipped)} SKUs moved in period ({basis})")
            firsts: list[tuple[str, str]] = []
            for batch in chunks(shipped, ID_BATCH):
                for r in mcp.rows(q_first(kinds, batch), "new items first-date"):
                    d = to_date(g(r, 2))
                    firsts.append((str(g(r, 1)), d.isoformat() if d else str(g(r, 2))))
            log(f"[newitems] {len(firsts)} first-fulfilled inside period")
            return firsts, basis
        except McpError as e:
            log(f"[newitems] {basis} query failed ({str(e)[:120]}) — "
                "trying fallback" if basis == "item_fulfillment" else
                f"[newitems] fallback also failed: {str(e)[:120]}")
            if basis == "first_sale_fallback":
                return [], f"unavailable: {str(e)[:160]}"
    return [], "unavailable"


def _dim_lookup(cust: dict | None, key: str) -> Any:
    """Safe read from the customer dimension — the row may have no match."""
    return cust.get(key) if cust else None



def id_num(v):
    """Doc numbers and item ids as NUMBERS when they are numeric — Mike's
    pasted tabs store them numeric, and the workbook's own VLOOKUP/MATCH
    against 'Commission Rate by SKUs' (and the H>threshold contracted-rate
    branches) are type-sensitive: our TEXT ids made VLOOKUP miss and the
    IFERROR default 10% override the contracted 7.5% on ~5,000 rows
    (found 0821, row-level diff of workbook AA vs the shadow engine)."""
    if v is None:
        return None
    s = str(v).strip()
    return int(s) if s.isdigit() else v

def build_ar_rows(raw: list[dict], cust: dict[str, dict],
                  items: dict[str, str], states: dict[str, str],
                  partners: dict[str, str], accounts: dict[str, str],
                  sign_flip: bool) -> list[list]:
    """24 values per row -> AR tab columns A:X.

    A Client:Project   B Customer_Type   C Client Category  D First Sale Date
    E Store_Type       F Date            G Transaction      H No.
    I Item: Full Name  J Item: Desc      K Quantity         L Open Balance
    M Sales Rep        N Address(state)  O Partner          P Partner Role
    Q Commission Pct   R Txn status      S Due Date         T Date Closed
    U Primary Partner  V Amount (Gross)  W Account: Name    X Company Name

    KNOWN SOURCE-WORKBOOK DEFECT — DO NOT "FIX" THIS MAPPING TO MATCH IT
    (found 2026-08-19 inspecting the hand-built 06.2026 file): AR_06.30 has
    a one-column shift on ~40% of rows — column Q holds the Transaction
    Status there (sharedString indices 1314="Open", 1468="Paid In Full",
    written into cells missing the t="s" attribute) and column R is only
    ~40% filled. THIS output is the correct layout: Q = Commission Pct
    (customer custentity12 label), R = status, 100% fill. Anyone
    reconciling Q or R against prior hand-built months will see a
    difference; that difference is the source's bug, not ours.
    """
    sf = -1 if sign_flip else 1
    out: list[list] = []
    skipped_partner = skipped_company = unmatched = 0

    for r in raw:
        partner = g(r, 9)
        if (partner or "None") in EXCLUDED_PARTNERS:
            skipped_partner += 1
            continue
        ent = str(g(r, 15))
        c = cust.get(ent)
        if not c:
            unmatched += 1
        company = _dim_lookup(c, "company")
        if company and company in EXCLUDED_COMPANIES:
            skipped_company += 1
            continue
        # AR is Romane-scoped, and only these store types are commissionable.
        if c and c.get("category") != "Romane":
            continue
        if c and c.get("store_type") not in STORE_TYPES:
            continue

        qty, bal, gross = num(g(r, 6)), num(g(r, 7)), num(g(r, 13))
        item_id = g(r, 5)
        out.append([
            g(r, 1),
            _dim_lookup(c, "type"),
            _dim_lookup(c, "category"),
            _dim_lookup(c, "first_sale"),
            _dim_lookup(c, "store_type"),
            serial(g(r, 2)),
            g(r, 3),
            id_num(g(r, 4)),
            id_num(item_id),
            items.get(str(item_id), item_id),
            qty * sf if isinstance(qty, float) else qty,
            bal * sf if isinstance(bal, float) else bal,
            employee_label(g(r, 8)),
            states.get(str(g(r, 14))),
            partner,
            partners.get(str(partner).strip(), PARTNER_ROLE_FALLBACK) if partner else None,
            _dim_lookup(c, "commission_pct"),       # Q: customer custentity12 / 100
            g(r, 10),
            serial(g(r, 11)),
            serial(g(r, 12)),
            _dim_lookup(c, "partner"),
            gross * sf if isinstance(gross, float) else gross,
            accounts.get(str(g(r, 14))),            # W: header AR account, DF'd
            company,
        ])

    return out, {"skipped_partner": skipped_partner,
                 "skipped_company": skipped_company,
                 "unmatched_customers": unmatched}


def build_sales_rows(raw: list[dict], cust: dict[str, dict],
                     items: dict[str, str], states: dict[str, str],
                     asof_serial: int, sign_flip: bool) -> list[list]:
    """25 values per row -> New Sales report / Sales report Raw columns A:Y.

    A:V match the AR layout EXCEPT P, and the W/X/Y tail differs — read off
    the delivered workbook 2026-08-18 (this used to be a copy of the AR
    shape, which is exactly how W/X/Y and P shipped wrong):

    L Amount           (same position AR calls "Open Balance")
    P Client: Partner  (the CUSTOMER's partner — a person — not the
                        partner-record category/role AR carries here)
    Q Commission Pct   (customer custentity12 / 100)
    W Month            ("Jun 2026" — the row's own date, %b %Y)
    X Client Age (Yrs) (years from firstsaledate to the period end, 2dp)
    Y Company name
    """
    sf = -1 if sign_flip else 1
    out: list[list] = []
    skipped_partner = skipped_company = unmatched = 0

    for r in raw:
        partner = g(r, 9)
        if (partner or "None") in EXCLUDED_PARTNERS:
            skipped_partner += 1
            continue
        ent = str(g(r, 15))
        c = cust.get(ent)
        if not c:
            unmatched += 1
        company = _dim_lookup(c, "company")
        if company and company in EXCLUDED_COMPANIES:
            skipped_company += 1
            continue

        qty, amt = num(g(r, 6)), num(g(r, 7))
        item_id = g(r, 5)
        first_sale = _dim_lookup(c, "first_sale")
        client_age = (round((asof_serial - first_sale) / 365.25, 2)
                      if isinstance(first_sale, (int, float)) else None)
        out.append([
            g(r, 1),
            _dim_lookup(c, "type"),
            _dim_lookup(c, "category"),
            first_sale,
            _dim_lookup(c, "store_type"),
            serial(g(r, 2)),
            g(r, 3),
            id_num(g(r, 4)),
            id_num(item_id),
            items.get(str(item_id), item_id),
            qty * sf if isinstance(qty, float) else qty,
            amt * sf if isinstance(amt, float) else amt,
            employee_label(g(r, 8)),
            states.get(str(g(r, 14))),
            partner,
            _dim_lookup(c, "partner"),              # P: Client: Partner (a person)
            _dim_lookup(c, "commission_pct"),       # Q: customer custentity12 / 100
            g(r, 10),
            serial(g(r, 11)),
            serial(g(r, 12)),
            _dim_lookup(c, "partner"),              # U: Primary Partner: Name
            amt * sf if isinstance(amt, float) else amt,
            month_label(g(r, 2)),                   # W: "Jun 2026"
            client_age,                             # X: Client Age (Years)
            company,                                # Y: Company name
        ])

    return out, {"skipped_partner": skipped_partner,
                 "skipped_company": skipped_company,
                 "unmatched_customers": unmatched}


# ─────────────────────────── formula columns ───────────────────────────
#
# The surgeon writes any string starting with '=' as a real formula, so there is
# no need for a fill_formulas op — the extract emits the formula TEXT per row.
#
# READ FROM THE REAL WORKBOOK 2026-08-17 via the surgeon's live formula probe
# (a job run against the actual 06.2026 source file on Render — the file is
# too large for any direct-read channel available here). Row 7 of AR_06.30 /
# New Sales report row 7 / Sales report Raw row 4.
#
# TWO cells (AR 'AD' and 'AG') hardcode the literal prior-tab name 'AR_05.31'
# (June's OWN prior month). Templating that literally would make every future
# month silently reference the wrong prior AR tab — August's run would still
# say 'AR_05.31' instead of 'AR_07.31'. Those two use a SECOND placeholder,
# {prior_ar_tab}, resolved per-run from the job's duplicate_sheet op (see
# commission_job.build_ops). Every other cell only substitutes {r}.
#
# NOT independently verified: the SEARCH()-matched retailer names (Cavenders,
# Scheels, Atwoods, The Glik Company, Quiet Storm, Bomgaar) and the numeric
# thresholds (H7>1121738, H7>1122227) are transcribed exactly as found — they
# are business rules (rate-change cutovers), not something this code can
# confirm is still current. Same for the 'Kevin Hanks' and 'Commission Rate
# by SKUs' sheet references (fixed lookup tabs, not month-dependent, so safe
# to carry literally). CONFIRM the first automated run's commission numbers
# by hand before trusting this unattended.

FORMULA_TEMPLATES = {
    "ar": {                       # AR tab, columns Y:AH
        "Y": ('=IF(AND(ISNUMBER(SEARCH("Cavenders",X{r})),H{r}>1121738,YEAR(F{r})<2026),0.09,'
              'IF(ISNUMBER(SEARCH("Cavenders",X{r})),0.0275,'
              'IF(AND(ISNUMBER(SEARCH("Scheels",X{r})),H{r}>1122227,YEAR(F{r})<2026),0.07,'
              'IF(OR(ISNUMBER(SEARCH("Atwoods",X{r})),ISNUMBER(SEARCH("Scheels",X{r})),'
              'ISNUMBER(SEARCH("The Glik Company",X{r})),ISNUMBER(SEARCH("Quiet Storm",X{r})),'
              'ISNUMBER(SEARCH("Glik",X{r}))),0.05,'
              'IF(AND(ISNUMBER(SEARCH("Bomgaar",X{r})),U{r}="Kelly Kennedy",AF{r}="Core",F{r}>$AI$1),0.06,'
              'IF(AND(ISNUMBER(SEARCH("Bomgaar",X{r})),U{r}="Kelly Kennedy",AF{r}="Licensed",F{r}>$AI$1),0.045,'
              '" "))))))'),
        "Z": "=V{r}-L{r}",        # CONFIRMED: Difference (Open Balance)
        "AA": ('=IFERROR(IF(OR(U{r}="8 Nolita",U{r}="None"),"0",'
               'IF(U{r}="Kevin Hanks",Y{r},'
               'IF(ISNUMBER(SEARCH("Bomgaars",X{r})),Y{r},'
               'IF(AND(U{r}=$AE$1,F{r}>$AD$1)," ",'
               'VLOOKUP(I{r},\'Commission Rate by SKUs\'!B:E,4,0))))),0.1)'),
        "AB": "=IFERROR(L{r}*AA{r},0)",
        "AC": '=+CONCATENATE(H{r}," - ",I{r})',
        # Lookup array is the prior tab's AG column — correct because of the
        # PRIOR-TAB REFRESH step (SOP, proven from the hand-built 06.2026
        # file): when a month rotates to "prior", its tab is re-pulled as-of
        # its month end WITH current payment info and converted to the
        # prior layout, where AG = "Concatenation No. & Item" (the key),
        # AE = unearned, AF = earned. Dashboard K/M/E and these lookups all
        # read that layout. v28 briefly pointed this at AC:AC — that was
        # only "right" for an UNREFRESHED prior tab, which is itself the
        # bug; with the refresh in place (extract_prior_ar), AG:AG is the
        # contract every month.
        "AD": ('=+IFERROR(IF(F{r}<$AD$4,_xlfn.XLOOKUP(AC{r},\'{prior_ar_tab}\'!AG:AG,'
               "'{prior_ar_tab}'!L:L),IF(F{r}>$AD$4,V{r},\" \")),0)"),
        "AE": ('=IF(AND(U{r}="Kevin Hanks",Y{r}=" ")," ",'
               'IF(U{r}=""," ",IF(AD{r}<>" ",L{r}-AD{r}," ")))'),
        "AF": '=IF(ISNUMBER(MATCH(I{r},\'Commission Rate by SKUs\'!$B:$B,0)),"Licensed","Core")',
        "AG": "=_xlfn.XLOOKUP(AC{r},'{prior_ar_tab}'!AG:AG,'{prior_ar_tab}'!L:L)",
        "AH": "=L{r}-AG{r}",
    },
    "sales": {                     # New Sales report, columns Z:AH
        "Z": "=_xlfn.XLOOKUP(H{r},'Kevin Hanks'!R:R,'Kevin Hanks'!S:S,\" \")",
        "AA": ('=IF(AND(ISNUMBER(SEARCH("Cavenders",Y{r})),H{r}>1121738,YEAR(F{r})<2026),0.09,'
               'IF(ISNUMBER(SEARCH("Cavenders",Y{r})),0.0275,'
               'IF(AND(ISNUMBER(SEARCH("Scheels",Y{r})),H{r}>1122227,YEAR(F{r})<2026),0.07,'
               'IF(OR(ISNUMBER(SEARCH("Atwoods",Y{r})),ISNUMBER(SEARCH("Scheels",Y{r})),'
               'ISNUMBER(SEARCH("The Glik Company",Y{r})),ISNUMBER(SEARCH("Quiet Storm",Y{r})),'
               'ISNUMBER(SEARCH("Glik",Y{r}))),0.05,'
               'IF(AND(ISNUMBER(SEARCH("Bomgaar",Y{r})),U{r}="Kelly Kennedy",AH{r}="Core",F{r}>$AI$1),0.06,'
               'IF(AND(ISNUMBER(SEARCH("Bomgaar",Y{r})),U{r}="Kelly Kennedy",AH{r}="Licensed",F{r}>$AI$1),0.045,'
               '" "))))))'),
        "AB": ('=_xlfn.IFS(AND(U{r}=$AI$4,T{r}>$AJ$4),"",G{r}="Cash Sale",F{r},'
               'ISBLANK(T{r}),"",T{r}<=$AC$6,T{r},T{r}>$AC$6,"")'),
        "AC": '=IFERROR(MONTH(AB{r}),"")',
        "AD": ('=IFERROR(IF(OR(U{r}="8 Nolita",U{r}="None"),"0",'
               'IF(U{r}="Kevin Hanks",AA{r},'
               'IF(ISNUMBER(SEARCH("Bomgaars",Y{r})),AA{r},'
               'IF(AND(U{r}=$AI$4,F{r}>$AH$4)," ",'
               'VLOOKUP(I{r},\'Commission Rate by SKUs\'!B:E,4,0))))),0.1)'),
        "AE": "=IFERROR(V{r}*AD{r},0)",
        "AF": ('=IFERROR(IF(OR(AC{r}*1<Dashboard!$C$34,AC{r}*1=Dashboard!$C$34),AE{r},""),"")'),
        "AG": '=_xlfn.CONCAT(H{r}," - ",I{r})',
        "AH": '=IF(ISNUMBER(MATCH(I{r},\'Commission Rate by SKUs\'!$B:$B,0)),"Licensed","Core")',
    },
    "raw": {                       # Sales report Raw, columns Z:AA
        "Z": '=TEXT(T{r},"MMM YY")',
        "AA": '=+CONCATENATE(H{r}," - ",I{r})',
    },
}


def col_to_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def formula_cells(kind: str, first_row: int, count: int,
                  prior_ar_tab: str | None = None) -> dict[str, str]:
    """{'Z7': '=V7-L7', 'Z8': ...} for a set_cells op, or fold into paste rows.

    prior_ar_tab fills the AR 'AD'/'AG' XLOOKUP templates' {prior_ar_tab}
    placeholder. Required whenever kind == 'ar' and those templates are
    populated — raises rather than silently emit a formula pointing at last
    month's prior tab."""
    out: dict[str, str] = {}
    for tpl_col, tpl in FORMULA_TEMPLATES.get(kind, {}).items():
        if "{prior_ar_tab}" in tpl and not prior_ar_tab:
            raise ValueError(
                f"{kind!r} template {tpl_col!r} needs prior_ar_tab but none was given")
        for i in range(count):
            r = first_row + i
            cell = tpl.replace("{r}", str(r))
            if prior_ar_tab:
                cell = cell.replace("{prior_ar_tab}", prior_ar_tab)
            out[f"{tpl_col}{r}"] = cell
    return out


# ─────────────────────────── top level ───────────────────────────

def extract(mcp: Mcp, asof: str, frm: str, to: str,
            sign_flip: bool = True, log=print) -> dict:
    log(f"[ar] querying open-at-{asof}")
    ar_raw = mcp.rows(q_ar(asof), "AR aging")
    log(f"[ar] {len(ar_raw)} raw lines")

    log(f"[sales] id sweep {frm}..{to}")
    ids = [str(g(r, 1)) for r in mcp.rows(q_sales_ids(frm, to), "sales ids")]
    log(f"[sales] {len(ids)} transactions")

    sales_raw: list[dict] = []
    for n, batch in enumerate(chunks(ids, ID_BATCH), 1):
        sales_raw.extend(mcp.rows(q_sales_lines(batch), f"sales lines {n}"))
        log(f"[sales] batch {n}: {len(sales_raw)} lines so far")

    ent_ids = sorted({str(g(r, 15)) for r in ar_raw + sales_raw if g(r, 15)})
    emp_codes = sorted({str(g(r, 8)) for r in ar_raw + sales_raw if g(r, 8)})
    item_ids = sorted({str(g(r, 5)) for r in ar_raw + sales_raw if g(r, 5)})
    unmapped = [c for c in emp_codes if c not in EMPLOYEE_NAMES]
    log(f"[dim] {len(ent_ids)} customers, {len(emp_codes)} employee codes "
        f"({len(unmapped)} unmapped), {len(item_ids)} items")
    if unmapped:
        log(f"[dim] add these to EMPLOYEE_NAMES: {', '.join(unmapped)}")

    txn_ids = sorted({str(g(r, 14)) for r in ar_raw + sales_raw if g(r, 14)})
    ar_txn_ids = sorted({str(g(r, 14)) for r in ar_raw if g(r, 14)})
    states = fetch_states(mcp, txn_ids)
    partners = fetch_partners(mcp)
    accounts = fetch_accounts(mcp, ar_txn_ids)   # AR col W (header AR account)
    log(f"[dim] {len(partners)} partner categories")
    log(f"[dim] {len(states)} shipping states")
    log(f"[dim] {len(accounts)} AR accounts")
    cust = fetch_customers(mcp, ent_ids)
    items: dict[str, str] = {}
    for batch in chunks(item_ids, CUST_BATCH):
        for r in mcp.rows(q_items(batch), "item dim"):
            desc = g(r, 2)
            if desc:
                items[str(g(r, 1))] = re.sub(r"\s+", " ", str(desc)).strip()

    asof_serial = serial(asof)
    # units-by-sku (new-items input) is the ONLY thing that still needs the
    # raw line dicts after the rows are built — compute it now so both raw
    # lists (30k dicts x 15 keys, 60-120MB) can be released immediately
    # after their build. The raw-stage OOMs kept happening with the extract
    # era's lifetime peak dominating the container.
    sf0 = -1 if sign_flip else 1
    units_by_sku: dict[str, float] = {}
    for r in sales_raw:
        q = num(g(r, 6))
        if isinstance(q, float):
            sku = str(g(r, 5))
            units_by_sku[sku] = units_by_sku.get(sku, 0.0) + q * sf0
    ar_rows, ar_diag = build_ar_rows(ar_raw, cust, items, states, partners,
                                     accounts, sign_flip)
    del ar_raw
    sales_rows, sales_diag = build_sales_rows(sales_raw, cust, items, states,
                                              asof_serial, sign_flip)
    del sales_raw

    # SOP step 4 input: SKUs first fulfilled inside the period, with units
    # sold this period (from the lines already extracted) and the real
    # product description. Consumers exclude SKUs already on the workbook's
    # 'Commission Rate by SKUs' tab (the service does this when readRanges
    # supplies that list).
    first_fulfilled, ni_basis = fetch_new_items(mcp, frm, to, log=log)
    # units_by_sku was computed above, before the raw lists were released
    ni_ids = [sku for sku, _d in first_fulfilled if sku not in items]
    for batch in chunks(ni_ids, CUST_BATCH):
        for r in mcp.rows(q_items(batch), "new item desc"):
            desc = g(r, 2)
            if desc:
                items[str(g(r, 1))] = re.sub(r"\s+", " ", str(desc)).strip()
    new_items = [{"sku": sku, "description": items.get(sku),
                  "firstFulfilled": d,
                  "unitsSold": round(units_by_sku.get(sku, 0.0), 2)}
                 for sku, d in first_fulfilled]

    ar_total = sum(r[11] for r in ar_rows if isinstance(r[11], float))
    sales_total = sum(r[11] for r in sales_rows if isinstance(r[11], float))

    return {
        "arRows": ar_rows, "salesRows": sales_rows,
        "arCount": len(ar_rows), "salesCount": len(sales_rows),
        "arOpenBalance": round(ar_total, 2),
        "salesAmount": round(sales_total, 2),
        "txnCount": len(ids),
        "newItems": new_items, "newItemsBasis": ni_basis,
        "diagnostics": {"ar": ar_diag, "sales": sales_diag,
                        "customers": len(cust), "employeeCodes": len(emp_codes),
                        "employeesUnmapped": unmapped,
                        "items": len(items), "states": len(states),
                        "partners": len(partners), "arAccounts": len(accounts),
                        "signFlip": sign_flip},
    }


# ──────────────────────────────────────────────────────────────────────
# PRIOR-TAB REFRESH — the SOP step the hand-built 06.2026 file proved.
#
# When a month rotates to "prior", Mike re-pulls its AR aging (same as-of
# date, but run DURING the following month, so `Date Closed` carries the
# payment dates that have happened since) and the tab takes the PRIOR
# LAYOUT (May row-6 headers): Y Contracted Rates, Z Difference,
# AA Deposit Month ('Jun26'), AB Client Age, AC Deposit Month(In number),
# AD Commission Rate, AE Unearned (=L*AD), AF Commission earned on
# receipts (=AE unless AC=$AB$4 future-month or 0 not-deposited),
# AG Concatenation No. & Item (the lookup key), AH left alone.
#
# Everything downstream depends on this refresh: Dashboard E's receipts
# term sums prior L where T (Date Closed) < the period cutoff; K sums
# prior AE; M sums prior AF; and the current tab's AD/AG XLOOKUPs key on
# prior AG. A FROZEN prior tab (what the pipeline shipped before this
# existed) leaves T stale — measured on 0820-1803: receipts mostly
# invisible, 12,506 column-shift junk values in T spuriously passing the
# cutoff, J total 594k vs Mike's reconciled 0.
# ──────────────────────────────────────────────────────────────────────

# The June-layout Y tests AF{r}="Core"/"Licensed" — fine there, where AF is
# the product type. In the PRIOR layout AF becomes EARNED, which chains
# AF -> AE -> AD -> Y: a circular reference (Excel warned on open,
# 0820 evening). Mike sidesteps it by flattening Y to values; we keep the
# formula but substitute the AF tests with the product-type check AF-as-type
# itself computes (MATCH against the SKU tab), which breaks the cycle with
# identical logic.
_LICENSED_TEST = "ISNUMBER(MATCH(I{r},'Commission Rate by SKUs'!$B:$B,0))"
_PRIOR_Y = (FORMULA_TEMPLATES["ar"]["Y"]
            .replace('AF{r}="Core"', f"NOT({_LICENSED_TEST})")
            .replace('AF{r}="Licensed"', _LICENSED_TEST))
assert 'AF{r}' not in _PRIOR_Y, "prior Y still references AF — cycle not broken"
assert _PRIOR_Y != FORMULA_TEMPLATES["ar"]["Y"], \
    "ar Y template changed shape — revisit the prior-layout substitution"

PRIOR_FORMULA_TEMPLATES: dict[str, str] = {
    "Y": _PRIOR_Y,                              # Contracted Rates, cycle-free
    "Z": "=V{r}-L{r}",                          # Difference
    "AD": FORMULA_TEMPLATES["ar"]["AA"],        # Commission Rate (same logic,
                                                # prior layout keeps it at AD)
    "AE": "=IFERROR(L{r}*AD{r},0)",             # Unearned
    "AF": '=IFERROR(IF(OR(AC{r}=$AB$4,AC{r}=0)," ",AE{r}),"0")',   # Earned
    "AG": '=+CONCATENATE(H{r}," - ",I{r})',     # the key the NEXT month looks up
}


def prior_formula_cells(first_row: int, count: int) -> dict[str, str]:
    out: dict[str, str] = {}
    for rn in range(first_row, first_row + count):
        for col, tpl in PRIOR_FORMULA_TEMPLATES.items():
            out[f"{col}{rn}"] = tpl.replace("{r}", str(rn))
    return out


def prior_value_cells(ar_rows: list[list], first_row: int,
                      asof_serial: int) -> dict[str, object]:
    """AA (deposit month label 'Jun26'), AB (client age in years),
    AC (deposit month number; 0 = not yet deposited) — plain VALUES, the
    way the hand-built prior tabs carry them. Derived from the row's own
    Date Closed (col T, index 19) and First Sale Date (col D, index 3)."""
    out: dict[str, object] = {}
    for i, row in enumerate(ar_rows):
        rn = first_row + i
        t = row[19]
        if isinstance(t, (int, float)):
            d = EXCEL_EPOCH + dt.timedelta(days=int(t))
            out[f"AA{rn}"] = f"{_MON3[d.month - 1]}{d.year % 100}"
            out[f"AC{rn}"] = d.month
        else:
            out[f"AC{rn}"] = 0
        first_sale = row[3]
        fs = serial(first_sale) if not isinstance(first_sale, (int, float)) \
            else first_sale
        if isinstance(fs, (int, float)):
            out[f"AB{rn}"] = round((asof_serial - fs) / 365, 2)
    return out


def norm_docno(v) -> str | None:
    """'726770', 726770.0 and ' 726770 ' must all key the same."""
    if v is None or v == "":
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def fetch_prior_closedates(mcp: Mcp, asof: str, log=print) -> dict[str, int]:
    """{document 'No.' -> Date Closed serial} for every invoice open at
    `asof`, evaluated NOW — the enrichment feed for the prior-tab refresh.

    Deliberately NOT a data re-pull: the tab's pasted L values (the as-of
    balances Mike closed the month on) stay untouched as ground truth.
    v29 replaced them with q_ar's reconstructed as-of balance and the
    total came back 1.02M against the tab's 1.78M — the postpay CTE
    under-reconstructs once weeks of payments separate the run from the
    as-of date. Mike's own refreshed tabs keep the historical balances
    (his May tab: closed rows still carry May-31 L) and only gain the
    close dates, so that's the contract."""
    log(f"[priorAr] pulling Date Closed for invoices open at {asof}")
    raw = mcp.rows(q_ar(asof), "prior AR close dates")
    out: dict[str, int] = {}
    for r in raw:
        key = norm_docno(g(r, 4))
        cs = serial(g(r, 12))
        if key and isinstance(cs, int):
            out[key] = cs
    log(f"[priorAr] {len(raw)} lines -> {len(out)} documents with a close date")
    return out


def enrich_prior_rows(rows: list[list], close_map: dict[str, int]) -> int:
    """Overwrite each existing row's T (col 20, 'Date Closed') with the
    VERIFIED close serial for its document number (col H), or clear it.
    Clearing is load-bearing: the hand-pasted June tab carries 12,506
    column-shift junk values in T (52, 60 — terms days) that spuriously
    pass Dashboard E's '<serial' receipt cutoff. After this, T holds only
    NetSuite-confirmed close dates. Returns the matched-row count."""
    hit = 0
    for row in rows:
        cs = close_map.get(norm_docno(row[7]) or "")
        row[19] = cs
        if cs is not None:
            hit += 1
    return hit


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--asof", required=True)
    ap.add_argument("--from", dest="frm", required=True)
    ap.add_argument("--to", required=True)
    ap.add_argument("--no-sign-flip", action="store_true")
    ap.add_argument("--url", default=os.getenv(
        "NS_MCP_URL", "https://netsuite-mcp-zdfd.onrender.com/mcp"))
    ap.add_argument("--secret", default=os.getenv("MCP_SHARED_SECRET", ""))
    ap.add_argument("--dump", help="write rows to this JSON file")
    a = ap.parse_args()

    if not a.secret:
        print("MCP_SHARED_SECRET not set (or pass --secret)", file=sys.stderr)
        return 2

    mcp = Mcp(a.url, a.secret)
    res = extract(mcp, a.asof, a.frm, a.to, sign_flip=not a.no_sign_flip)

    print("\n" + "=" * 58)
    print(f"AR rows       {res['arCount']:>12,}   target 16,362")
    print(f"AR open bal   {res['arOpenBalance']:>12,.2f}   target 2,680,818.23")
    print(f"Transactions  {res['txnCount']:>12,}   target ~5,165")
    print(f"Sales lines   {res['salesCount']:>12,}   target ~13,500")
    print(f"Sales amount  {res['salesAmount']:>12,.2f}")
    print("=" * 58)
    print(json.dumps(res["diagnostics"], indent=2))

    if res["arRows"]:
        print("\nfirst AR row (24 values):")
        print(json.dumps(res["arRows"][0], indent=2, default=str))
    if res["salesRows"]:
        print("\nfirst sales row (25 values):")
        print(json.dumps(res["salesRows"][0], indent=2, default=str))

    if a.dump:
        with open(a.dump, "w") as f:
            json.dump({k: v for k, v in res.items()}, f)
        print(f"\nwrote {a.dump}")

    ok = res["arCount"] > 0 and res["salesCount"] > 0
    print("\n" + ("LOOKS SANE" if ok else "EMPTY RESULT — do not proceed"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
