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

def serial(v: Any) -> Any:
    """Excel serial date. NetSuite hands back M/D/YYYY or ISO."""
    if v in (None, ""):
        return None
    s = str(v)
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        d = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    else:
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
        if not m:
            return v
        d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return (d - EXCEL_EPOCH).days


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
    """Rule 2: transaction + transactionline + item ONLY."""
    return f"""
SELECT
  BUILTIN.DF(t.entity) AS c01, t.trandate AS c02, BUILTIN.DF(t.type) AS c03,
  t.tranid AS c04, i.itemid AS c05, tl.quantity AS c06, tl.foreignamount AS c07,
  BUILTIN.DF(t.employee) AS c08, BUILTIN.DF(t.partner) AS c09,
  BUILTIN.DF(t.status) AS c10, t.duedate AS c11, t.closedate AS c12,
  tl.rate AS c13, t.id AS c14, t.entity AS c15
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

    c07 is the customer's shipping state. Column N of the workbook is
    customer-level (values are constant per client), which is why it belongs
    here and NOT as a transactionShippingAddress join — see rule 2.
    """
    return f"""
SELECT cu.id AS c01,
  BUILTIN.DF(cu.custentitycustomer_type) AS c02,
  BUILTIN.DF(cu.category) AS c03,
  cu.firstsaledate AS c04,
  BUILTIN.DF(cu.custentitystore_type) AS c05,
  cu.companyname AS c06,
  cu.custentity_shipping_state AS c07,      -- VERIFY: field name unconfirmed
  BUILTIN.DF(cu.partner) AS c08,
  cu.partnercategory AS c09                 -- VERIFY: 'Primary Rep' role, col P
FROM customer cu
WHERE cu.id IN ({','.join(ids)})
ORDER BY cu.id
""".strip()


def q_items(ids: list[str]) -> str:
    """Real product descriptions for column J.
    `item.salesdescription` returns HTTP 500 on this account; `description` is
    the fallback. If both fail, fall back to itemid and flag it."""
    return f"""
SELECT i.id AS c01, i.itemid AS c02, i.description AS c03
FROM item i
WHERE i.id IN ({','.join(ids)})
""".strip()


def q_employees(ids: list[str]) -> str:
    """Column M renders as 'EMP101 Mark Saviski' — code AND name.
    BUILTIN.DF(t.employee) yields only 'EMP101', so join the name here."""
    return f"""
SELECT e.id AS c01, e.entityid AS c02, e.firstname AS c03, e.lastname AS c04
FROM employee e
WHERE e.id IN ({','.join(ids)})
""".strip()


# ─────────────────────────── dimensions ───────────────────────────

def fetch_customers(mcp: Mcp, ids: list[str]) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for batch in chunks(ids, CUST_BATCH):
        for r in mcp.rows(q_customers(batch), "customer dim"):
            out[str(g(r, 1))] = {
                "type": g(r, 2), "category": g(r, 3),
                "first_sale": serial(g(r, 4)), "store_type": g(r, 5),
                "company": g(r, 6), "state": g(r, 7),
                "partner": g(r, 8), "partner_role": g(r, 9),
            }
    return out


def fetch_employees(mcp: Mcp, ids: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for batch in chunks(ids, CUST_BATCH):
        for r in mcp.rows(q_employees(batch), "employee dim"):
            code = g(r, 2) or ""
            name = " ".join(x for x in (g(r, 3), g(r, 4)) if x)
            out[str(g(r, 1))] = f"{code} {name}".strip()
    return out


# ─────────────────────────── assembly ───────────────────────────

def _dim_lookup(cust: dict, key: str) -> Any:
    return cust.get(key) if cust else None


def build_ar_rows(raw: list[dict], cust: dict[str, dict], emp: dict[str, str],
                  items: dict[str, str], sign_flip: bool) -> list[list]:
    """24 values per row -> AR tab columns A:X.

    A Client:Project   B Customer_Type   C Client Category  D First Sale Date
    E Store_Type       F Date            G Transaction      H No.
    I Item: Full Name  J Item: Desc      K Quantity         L Open Balance
    M Sales Rep        N Address(state)  O Partner          P Partner Role
    Q Commission %     R Txn status      S Due Date         T Date Closed
    U Primary Partner  V Amount (Gross)  W Account: Name    X Company Name
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
            g(r, 4),
            item_id,
            items.get(str(item_id), item_id),
            qty * sf if isinstance(qty, float) else qty,
            bal * sf if isinstance(bal, float) else bal,
            emp.get(str(g(r, 8)), g(r, 8)),
            _dim_lookup(c, "state"),
            partner,
            _dim_lookup(c, "partner_role"),
            None,                                   # rule 6: commission % empty
            g(r, 10),
            serial(g(r, 11)),
            serial(g(r, 12)),
            _dim_lookup(c, "partner"),
            gross * sf if isinstance(gross, float) else gross,
            None,                                   # W Account: Name (GL) — formula
            company,
        ])

    return out, {"skipped_partner": skipped_partner,
                 "skipped_company": skipped_company,
                 "unmatched_customers": unmatched}


def build_sales_rows(raw: list[dict], cust: dict[str, dict], emp: dict[str, str],
                     items: dict[str, str], sign_flip: bool) -> list[list]:
    """25 values per row -> New Sales report / Sales report Raw columns A:Y.
    Same A:X shape as AR, plus Y = rate."""
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

        qty, amt, rate = num(g(r, 6)), num(g(r, 7)), num(g(r, 13))
        item_id = g(r, 5)
        out.append([
            g(r, 1),
            _dim_lookup(c, "type"),
            _dim_lookup(c, "category"),
            _dim_lookup(c, "first_sale"),
            _dim_lookup(c, "store_type"),
            serial(g(r, 2)),
            g(r, 3),
            g(r, 4),
            item_id,
            items.get(str(item_id), item_id),
            qty * sf if isinstance(qty, float) else qty,
            amt * sf if isinstance(amt, float) else amt,
            emp.get(str(g(r, 8)), g(r, 8)),
            _dim_lookup(c, "state"),
            partner,
            _dim_lookup(c, "partner_role"),
            None,
            g(r, 10),
            serial(g(r, 11)),
            serial(g(r, 12)),
            _dim_lookup(c, "partner"),
            amt * sf if isinstance(amt, float) else amt,
            None,
            company,
            rate,
        ])

    return out, {"skipped_partner": skipped_partner,
                 "skipped_company": skipped_company,
                 "unmatched_customers": unmatched}


# ─────────────────────────── formula columns ───────────────────────────
#
# The surgeon writes any string starting with '=' as a real formula, so there is
# no need for a fill_formulas op — the extract emits the formula TEXT per row.
#
# These templates MUST be read off the prior month's sheet, not invented. Only
# the AR 'Z' column is confirmed (=V{r}-L{r}, seen at Z918 in 06.2026). Fill the
# rest by clicking row 7 of AR_06.30 / New Sales report / Sales report Raw.
#
# A row's formula cells are produced by substituting {r} with the sheet row.

FORMULA_TEMPLATES = {
    "ar": {                       # AR tab, columns Y:AH
        "Z": "=V{r}-L{r}",        # CONFIRMED: Difference (Open Balance)
        # "Y":  "=...",           # Contracted
        # "AA": "=...",           # Commission Rate
        # "AB": "=...",           # <Month> Unearned Commission
        # "AC": "=...",           # Concatenation No. & Item
        # "AD": "=...",           # Last month open balance
        # "AE": "=...",           # Partial Payment
        # "AF": "=...",           # Product Type
        # "AG"/"AH": Delete later — probably safe to leave empty
    },
    "sales": {},                  # New Sales report, columns Z:AH
    "raw": {},                    # Sales report Raw, columns Z:AA
}


def col_to_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def formula_cells(kind: str, first_row: int, count: int) -> dict[str, str]:
    """{'Z7': '=V7-L7', 'Z8': ...} for a set_cells op, or fold into paste rows."""
    out: dict[str, str] = {}
    for tpl_col, tpl in FORMULA_TEMPLATES.get(kind, {}).items():
        for i in range(count):
            r = first_row + i
            out[f"{tpl_col}{r}"] = tpl.replace("{r}", str(r))
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
    emp_ids = sorted({str(g(r, 8)) for r in ar_raw + sales_raw if g(r, 8)})
    item_ids = sorted({str(g(r, 5)) for r in ar_raw + sales_raw if g(r, 5)})
    log(f"[dim] {len(ent_ids)} customers, {len(emp_ids)} employees, "
        f"{len(item_ids)} items")

    cust = fetch_customers(mcp, ent_ids)
    emp = fetch_employees(mcp, emp_ids)
    items: dict[str, str] = {}
    for batch in chunks(item_ids, CUST_BATCH):
        for r in mcp.rows(q_items(batch), "item dim"):
            items[str(g(r, 1))] = g(r, 3) or g(r, 2)

    ar_rows, ar_diag = build_ar_rows(ar_raw, cust, emp, items, sign_flip)
    sales_rows, sales_diag = build_sales_rows(sales_raw, cust, emp, items, sign_flip)

    ar_total = sum(r[11] for r in ar_rows if isinstance(r[11], float))
    sales_total = sum(r[11] for r in sales_rows if isinstance(r[11], float))

    return {
        "arRows": ar_rows, "salesRows": sales_rows,
        "arCount": len(ar_rows), "salesCount": len(sales_rows),
        "arOpenBalance": round(ar_total, 2),
        "salesAmount": round(sales_total, 2),
        "txnCount": len(ids),
        "diagnostics": {"ar": ar_diag, "sales": sales_diag,
                        "customers": len(cust), "employees": len(emp),
                        "items": len(items), "signFlip": sign_flip},
    }


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
