"""Credit-memo receipts adjustment (Preet, 2026-09-10).

Dashboard J's usual variance: a credit memo dated in the period but applied
to a prior-period invoice. The hand-built 08.2026 book carries Mike's fix
as trailing constants on the reps' E formulas — `+H8-24` (Kevin, CM 180346
closed 7/30) and `+H21-17.5` (Kelly, CM 180351 closed in June). This
reproduces exactly those constants from the extracted sales rows, and
proves next month's existing strip regex removes them again.
"""
import sys

sys.path.insert(0, r"C:\DanielMirochnik\nonexistent")  # keep path list shape stable
sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from commission_job import credit_memo_adjustments
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:300]}"))
    if not cond:
        fails.append(label)


def row(gtype, doc, item, closed, partner, amount):
    r = [None] * 25
    r[6], r[7], r[8], r[19], r[20], r[21] = gtype, doc, item, closed, partner, amount
    return r


AUG1 = 46235
SALES = [
    # Kevin: CM dated in Aug, closed 7/30 -> -24 (three lines in Preet's book net to -24)
    row("Credit Memo", "180346", "98083", 46233, "Kevin Hanks", -24.0),
    row("Credit Memo", "180346", "98083", 46233, "Kevin Hanks", -5.89),
    row("Credit Memo", "180346", "98083", 46233, "Kevin Hanks", 5.89),
    # Kelly: CM closed in June -> -17.5
    row("Credit Memo", "180351", "92701", 46181, "Kelly Kennedy", -17.5),
    # CM closed INSIDE the period: a normal receipt, not an adjustment
    row("Credit Memo", "180400", "90092", 46240, "Brian Haberman", -50.0),
    # CM still open (no Date Closed): stays in AR, not an adjustment
    row("Credit Memo", "180401", "90092", None, "Brian Haberman", -60.0),
    # CM closed before the period for a partner NOT on the Dashboard table
    row("Credit Memo", "180402", "90092", 46200, "Somebody Else", -9.0),
    # invoices never qualify
    row("Invoice", "1163757", "98083", 46268, "Kevin Hanks", 3750.0),
    row("Invoice", "1160000", "98083", 46220, "Kevin Hanks", 100.0),
    # a boolean in T must not be mistaken for a serial
    row("Credit Memo", "180403", "90092", True, "Kevin Hanks", -1.0),
]
PARTNER_ROWS = {"Tiffany McDaniel": 6, "Walter Namiotka": 7, "Kevin Hanks": 8,
                "Dayna Stambeck": 12, "Brian Haberman": 13, "Kelly Kennedy": 21}

ops, report = credit_memo_adjustments(SALES, AUG1, PARTNER_ROWS)
check("A1 one replace_formula_text op on the Dashboard",
      len(ops) == 1 and ops[0]["op"] == "replace_formula_text" and ops[0]["sheet"] == "Dashboard", ops)
maps = {m["from"]: m for m in ops[0]["replace"]} if ops else {}
check("A2 Kevin: (\\+H8)$ -> \\g<1>-24.00 (the three CM lines net to -24)",
      maps.get(r"(\+H8)$", {}).get("to") == r"\g<1>-24.00", maps)
check("A3 Kelly: (\\+H21)$ -> \\g<1>-17.50",
      maps.get(r"(\+H21)$", {}).get("to") == r"\g<1>-17.50", maps)
check("A4 in-period and still-open credit memos do not adjust (Brian absent)",
      not any("H13" in k for k in maps), maps)
check("A5 invoices and boolean T ignored", len(maps) == 2, maps)
check("A6 unmatched partner reported, not silently dropped",
      report["unmatchedPartners"] and report["unmatchedPartners"][0]["partner"] == "Somebody Else"
      and report["unmatchedPartners"][0]["amount"] == -9.0, report)
check("A7 mappings are regex + optional (a blank E row must not fail the job)",
      all(m.get("regex") and m.get("optional") for m in ops[0]["replace"]), ops)
check("A8 applied report carries row, suffix and the source documents",
      any(a["partner"] == "Kevin Hanks" and a["row"] == 8 and a["formulaSuffix"] == "-24.00"
          and {i["doc"] for i in a["items"]} == {"180346"} for a in report["applied"]), report)

ops0, rep0 = credit_memo_adjustments([row("Invoice", "1", "2", 46240, "Kevin Hanks", 10.0)], AUG1, PARTNER_ROWS)
check("A9 no qualifying credit memos -> no op at all", ops0 == [] and rep0["applied"] == [], (ops0, rep0))

# ---- apply through the surgeon's formula-text engine on a Dashboard-shaped fragment
s = XlsxSurgeon.__new__(XlsxSurgeon)
xml = ('<c r="E8"><f>SUMIFS(\'AR_07.31\'!$L:$L,\'AR_07.31\'!$U:$U,Dashboard!B8,\'AR_07.31\'!$T:$T,"&lt;46266")'
       '+SUMIFS(\'New Sales report\'!$V:$V,\'New Sales report\'!$U:$U,Dashboard!B8)+H8-7897.5</f><v>1</v></c>'
       '<c r="E21"><f>SUMIFS(\'AR_07.31\'!$L:$L,\'AR_07.31\'!$U:$U,Dashboard!B21)+H21</f><v>2</v></c>'
       '<c r="E13"><f>SUMIFS(\'AR_07.31\'!$L:$L,\'AR_07.31\'!$U:$U,Dashboard!B13)+H13</f><v>3</v></c>'
       '<c r="W8"><f>+M8+N8+O8</f><v>4</v></c>'
       '<c r="E38"><f>+H27</f><v>5</v></c>'
       '<c r="R8"><f>+K8+L8-M8</f><v>6</v></c>')
STRIP = {"from": r"(\+H\d+)(?:[-+][0-9]+(?:\.[0-9]+)?)+$", "to": r"\1", "regex": True, "optional": True}
# same order as the real job: caller's strip of last month's constants, then ours
out1, c1 = s._apply_freplace(xml, [STRIP])
out2, c2 = s._apply_freplace(out1, ops[0]["replace"])
check("F1 last month's constant stripped first (+H8-7897.5 -> +H8)",
      "+H8-7897.5" not in out1 and "Dashboard!B8)+H8</f>" in out1, out1[:300])
check("F2 Kevin's E8 now ends +H8-24.00", "Dashboard!B8)+H8-24.00</f>" in out2, out2)
check("F3 Kelly's E21 now ends +H21-17.50", "+H21-17.50</f>" in out2, out2)
check("F4 Brian's E13, W8, R8 and E38 (+H27) untouched",
      "+H13</f>" in out2 and "<f>+M8+N8+O8</f>" in out2 and "<f>+K8+L8-M8</f>" in out2
      and "<f>+H27</f>" in out2, out2)
check("F5 exactly one hit per mapping", c2 == {r"(\+H8)$": 1, r"(\+H21)$": 1}, c2)
out3, c3 = s._apply_freplace(out2, [STRIP])
check("F6 next month's strip removes our constants again (one-book lifecycle)",
      "Dashboard!B8)+H8</f>" in out3 and "+H21</f>" in out3 and "-24.00" not in out3 and "-17.50" not in out3, out3)

print("\n" + ("ALL PASS" if not fails else f"{len(fails)} FAIL: {fails}"))
sys.exit(1 if fails else 0)
