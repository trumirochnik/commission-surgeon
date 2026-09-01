"""Prior-tab close-date enrichment (Mike/Preet review 2026-08-25).

The old pull queried q_ar(asof) — invoices still OPEN at the as-of date —
so any document whose closedate landed ON or BEFORE month end was invisible
and its Date Closed got cleared. All 12 documents in Mike's review list
have closedate <= asof; live NetSuite also shows tranid COLLISIONS across
transaction types (180249 = 2014 invoice AND 2026 credit memo), so the map
must be type-aware with a doc-only fallback.
"""
import sys

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from netsuite_extract import fetch_prior_closedates, enrich_prior_rows

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:260]}"))
    if not cond:
        fails.append(label)


class FakeMcp:
    def __init__(self, canned):
        self.canned = canned
        self.queries = []

    def rows(self, q, label):
        self.queries.append(q)
        # return only the canned rows whose tranid appears in this batch
        return [r for r in self.canned if f"'{r['c01']}'" in q]


CANNED = [
    # Mike's example: CM closed 6/19/2026 (serial 46193), BEFORE asof 6/30
    {"c01": "180249", "c02": "Credit Memo", "c03": "6/19/2026"},
    # same doc number, different type, later close — must not collide
    {"c01": "180249", "c02": "Invoice", "c03": "7/2/2026"},
    # plain invoice closed after asof (the only case the OLD query covered)
    {"c01": "726770", "c02": "Invoice", "c03": "7/15/2026"},
    # duplicate (type, doc) rows: LATEST close must win
    {"c01": "555001", "c02": "Invoice", "c03": "5/1/2026"},
    {"c01": "555001", "c02": "Invoice", "c03": "6/10/2026"},
]

mcp = FakeMcp(CANNED)
docnos = {"180249", "726770", "555001", "999999"}
typed, by_doc = fetch_prior_closedates(mcp, docnos, log=lambda *a: None)

q0 = mcp.queries[0]
check("Q1: query has NO open-at-asof filter",
      "closedate >" not in q0 and "closedate IS NULL" not in q0
      and "postpay" not in q0, q0)
check("Q2: query keys on the tab's document numbers",
      "t.tranid IN (" in q0 and "'180249'" in q0, q0)
check("Q3: only closed docs requested", "closedate IS NOT NULL" in q0)

check("M1: typed map disambiguates the 180249 collision",
      typed[("credit memo", "180249")] != typed[("invoice", "180249")], typed)
check("M2: latest close wins on duplicate (type, doc)",
      typed[("invoice", "555001")] == by_doc["555001"]
      and by_doc["555001"] == typed[("invoice", "555001")]
      and typed[("invoice", "555001")] > 46100, typed)
check("M3: unknown docs simply absent", "999999" not in by_doc)

# batching: 900 docnos -> 3 queries
mcp2 = FakeMcp([])
fetch_prior_closedates(mcp2, {f"d{i}" for i in range(900)}, log=lambda *a: None)
check("B1: 900 docnos batch into 3 queries", len(mcp2.queries) == 3,
      len(mcp2.queries))

# ── enrichment ──
def tab_row(ttype, doc, t_junk):
    r = [None] * 24
    r[6] = ttype
    r[7] = doc
    r[19] = t_junk
    return r

rows = [
    tab_row("Credit Memo", "180249", 52),      # junk T; typed hit -> 6/19
    tab_row("Invoice", "180249", None),        # same doc, invoice -> 7/2
    tab_row("CM", "726770", 60),               # unknown type label -> doc fallback
    tab_row("Invoice", "999999", 45),          # no close anywhere -> CLEARED
]
hit = enrich_prior_rows(rows, typed, by_doc)
check("E1: typed match lands per transaction type",
      rows[0][19] == typed[("credit memo", "180249")]
      and rows[1][19] == typed[("invoice", "180249")]
      and rows[0][19] != rows[1][19], [rows[0][19], rows[1][19]])
check("E2: doc-only fallback when the tab's type label is odd",
      rows[2][19] == by_doc["726770"], rows[2][19])
check("E3: junk T cleared when nothing verifies", rows[3][19] is None)
check("E4: hit count", hit == 3, hit)

print()
print("ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
