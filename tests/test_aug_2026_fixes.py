"""Aug 2026 parallel-run fixes (2026-09-10).

The first real month-end comparison against Preet's hand-built 08.2026 book
found every wrong number traced to two record types the integration role
could not View (Cash Sale, Customer Payment) plus one append that did not
de-duplicate. NetSuite returns an EMPTY set, not an error, for invisible
record types — so all three shipped silently. This locks in:

  * q_ar's payment-date fallback (record-id boundary when the payment
    record is invisible; real trandate whenever it is visible)
  * the pre-write gates: zero Cash Sales for the period -> refuse;
    invoices closed after month-end but a zero add-back -> refuse
  * the AR history window covering the 2022-09-07 open credit memo
  * trim_rows: the streaming pre-append trim that makes the cumulative
    'Sales report Raw' append idempotent
  * build_ops emitting trim_rows before append_rows for the raw tab
"""
import os, re, sys, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
import netsuite_extract as ne
from surgeon import XlsxSurgeon
from commission_job import build_ops

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:300]}"))
    if not cond:
        fails.append(label)


# ---------------------------------------------------------------- queries
q = ne.q_ar("2026-08-31", 29910297)
check("Q1 proxy clause: invisible payment links classified by id boundary",
      "ntll.nextdoc > 29910297" in q and "pt.trandate IS NULL" in q
      and "LEFT JOIN transaction pt" in q, q[:400])
check("Q2 real trandate still decides when the record IS visible",
      "pt.trandate > TO_DATE('2026-08-31','YYYY-MM-DD')" in q)
check("Q3 add-back limited to balance-reducing link types",
      "'CustPymt'" in q and "'Journal'" in q and "'CustCred'" in q)
q0 = ne.q_ar("2026-08-31")
check("Q4 legacy shape preserved when no boundary is given",
      "LEFT JOIN transaction pt" not in q0 and "nextdoc >" not in q0
      and "JOIN transaction pt ON pt.id = ntll.nextdoc" in q0
      and "pt.trandate IS NULL" not in q0, q0[:300])
check("Q5 next_month_start incl. December rollover",
      ne.next_month_start("2026-08-31") == "2026-09-01"
      and ne.next_month_start("2026-12-31") == "2027-01-01")
check("Q6 boundary query keys on createddate before next month",
      "createddate < TO_DATE('2026-09-01'" in ne.q_boundary_id("2026-09-01"))
check("H1 AR history window covers the 2022-09-07 open credit memo",
      ne.AR_HISTORY_START <= "2022-09-07", ne.AR_HISTORY_START)
# first v48 run (2026-09-10) died with "offset out of bounds (0..99000)": the
# role could suddenly see cash sales, which never close, so q_ar admitted
# 27,705 of them (87,430 lines). AR aging = invoices + credit memos only.
_ar_where = q[q.index("FROM transaction t"):]
check("Q7 AR aging pull excludes Cash Sales (they never get a closedate)",
      "'CashSale'" not in _ar_where and "'CustInvc'" in _ar_where
      and "'CustCred'" in _ar_where, _ar_where[:400])
check("Q8 the SALES sweep still includes Cash Sales",
      "'CashSale'" in ne.q_sales_ids("2026-08-01", "2026-08-31"))


# ---------------------------------------------------------------- gates
class FakeMcp:
    """Answers the pre-write gate queries; anything past them raises a
    sentinel so a HEALTHY extract is recognisable without faking the
    dimension pulls."""

    def __init__(self, vis, ar_rows, postpay):
        self.vis, self.ar, self.pp = vis, ar_rows, postpay
        self.labels = []

    def rows(self, q, label=""):
        self.labels.append(label)
        if "MAX(t.id)" in q:
            return [{"c01": "29910297"}]
        if "GROUP BY t.type" in q:
            return [{"c01": k, "c02": str(v)} for k, v in self.vis.items()]
        if "WITH postpay" in q and "COUNT(*) AS c01" in q:
            return [{"c01": str(self.pp[0]), "c02": str(self.pp[1])}]
        if "WITH postpay" in q:
            return self.ar
        raise RuntimeError("SENTINEL past the gates: " + label)


NOLOG = lambda *a, **k: None
AR_CLOSED_AFTER = [{"c12": "9/3/2026", "c14": "29812935"}]

blind = FakeMcp({"CustInvc": 6359, "CustCred": 38}, AR_CLOSED_AFTER, (0, 0))
try:
    ne.extract(blind, "2026-08-31", "2026-08-01", "2026-08-31", log=NOLOG)
    check("G1 refuses when the role sees no Cash Sales", False, "no exception")
except ValueError as e:
    check("G1 refuses when the role sees no Cash Sales", "Cash Sale" in str(e), e)
except Exception as e:  # noqa: BLE001
    check("G1 refuses when the role sees no Cash Sales", False, repr(e))

nopay = FakeMcp({"CustInvc": 6359, "CashSale": 188, "CustCred": 38}, AR_CLOSED_AFTER, (0, 0))
try:
    ne.extract(nopay, "2026-08-31", "2026-08-01", "2026-08-31", log=NOLOG)
    check("G2 refuses when invoices closed after month-end but add-back is zero", False, "no exception")
except ValueError as e:
    check("G2 refuses when invoices closed after month-end but add-back is zero",
          "add-back" in str(e), e)
except Exception as e:  # noqa: BLE001
    check("G2 refuses when invoices closed after month-end but add-back is zero", False, repr(e))
check("G2b the AR query used the boundary proxy (no CustPymt visibility)",
      any("AR aging" in l for l in nopay.labels))

healthy = FakeMcp({"CustInvc": 6359, "CashSale": 188, "CustCred": 38},
                  AR_CLOSED_AFTER, (256, 495751.67))
try:
    ne.extract(healthy, "2026-08-31", "2026-08-01", "2026-08-31", log=NOLOG)
    check("G3 healthy extract passes the gates", False, "no exception")
except RuntimeError as e:
    check("G3 healthy extract passes the gates and proceeds to the sales sweep",
          "SENTINEL" in str(e) and "sales ids" in str(e), e)
except Exception as e:  # noqa: BLE001
    check("G3 healthy extract passes the gates and proceeds to the sales sweep", False, repr(e))

# no invoice closed after as-of and no add-back: legitimate, must NOT refuse
quiet = FakeMcp({"CustInvc": 6359, "CashSale": 188, "CustCred": 38}, [], (0, 0))
try:
    ne.extract(quiet, "2026-08-31", "2026-08-01", "2026-08-31", log=NOLOG)
    check("G4 zero add-back with nothing closed after as-of is not a refusal", False, "no exception")
except RuntimeError as e:
    check("G4 zero add-back with nothing closed after as-of is not a refusal", "SENTINEL" in str(e), e)
except Exception as e:  # noqa: BLE001
    check("G4 zero add-back with nothing closed after as-of is not a refusal", False, repr(e))


# ---------------------------------------------------------------- trim_rows
TD = tempfile.mkdtemp(prefix="trim_")
src_part = os.path.join(TD, "raw_src.xml")
# rows 1-3 header (row 3 carries a "date" that must NOT be trimmed), data 4..13:
#   F serials 46200.. ; rows 10-13 are >= 46235 (period start) -> dropped
#   row 8: F is a shared string (t="s") -> kept ; row 9: no F cell -> kept
rows_xml = ['<row r="1"><c r="A1" t="inlineStr"><is><t>Tru</t></is></c></row>',
            '<row r="2"/>',
            '<row r="3"><c r="A3" t="inlineStr"><is><t>Client</t></is></c>'
            '<c r="F3"><v>99999</v></c></row>']
for r in range(4, 14):
    if r == 8:
        fcell = f'<c r="F{r}" t="s"><v>12</v></c>'
    elif r == 9:
        fcell = ""
    else:
        fcell = f'<c r="F{r}" s="5"><v>{46200 + (r - 4) * 6}</v></c>'   # 46200..46254
    rows_xml.append(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>C{r}</t></is></c>'
                    f'{fcell}<c r="L{r}"><v>{r * 10}</v></c></row>')
with open(src_part, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<dimension ref="A1:L13"/><sheetData>' + "".join(rows_xml)
            + '</sheetData><autoFilter ref="A3:L13"/></worksheet>')

s = XlsxSurgeon.__new__(XlsxSurgeon)
s.workdir = TD
dst_part = os.path.join(TD, "raw_out.xml")
new_rows = [["N1", "", "", "", "", 46237, "Cash Sale", "147600", "98083", "", 1, 600.0],
            ["N2", "", "", "", "", 46265, "Invoice", "1169061", "90092", "", 2, 135.0]]
changed = s._transform_sheet(src_part, dst_part, {}, [(new_rows, None)],
                             trim_groups=[("F", 46235.0, 4)])
out = open(dst_part, encoding="utf-8").read()
kept = sorted(int(x) for x in re.findall(r'<row r="(\d+)"', out))


def has_text(xml, txt):
    return re.search(r"<t[^>]*>" + re.escape(txt) + r"</t>", xml) is not None


check("T1 rows dated on/after the period start were dropped (their content is gone)",
      not any(has_text(out, f"C{r}") for r in (10, 11, 12, 13)), kept)
check("T2 earlier data rows kept, incl. shared-string F (8) and missing F (9)",
      all(has_text(out, f"C{r}") for r in (4, 5, 6, 7, 8, 9)), kept)
check("T3 header rows untouched even though row 3 holds a large number",
      all(r in kept for r in (1, 2, 3)) and '<c r="F3"><v>99999</v></c>' in out, kept)
check("T4 appended rows continue right after the last kept row (r=10, r=11)",
      kept == list(range(1, 12))
      and re.search(r'<row r="10"[^>]*>.*?<t[^>]*>N1</t>', out, re.S)
      and re.search(r'<row r="11"[^>]*>.*?<t[^>]*>N2</t>', out, re.S), kept)
check("T5 autoFilter re-pinned to the new last row",
      '<autoFilter ref="A3:L11"/>' in out,
      re.findall(r'<(?:dimension|autoFilter) ref="[^"]*"', out))
check("T6 change count includes the 4 dropped rows + appended cells",
      changed == 4 + sum(len(r) for r in new_rows), changed)
check("T7 trimmed-row count exposed for opsResults",
      getattr(s, "_last_trimmed", None) == 4, getattr(s, "_last_trimmed", None))

# idempotence: running the same trim+append on the OUTPUT yields the same rows
dst2 = os.path.join(TD, "raw_out2.xml")
s._transform_sheet(dst_part, dst2, {}, [(new_rows, None)], trim_groups=[("F", 46235.0, 4)])
out2 = open(dst2, encoding="utf-8").read()
check("T8 trim+append is idempotent (second pass reproduces the first)",
      re.findall(r'<row r="(\d+)"', out2) == re.findall(r'<row r="(\d+)"', out)
      and len(re.findall(r"<t[^>]*>N1</t>", out2)) == 1
      and len(re.findall(r"<t[^>]*>N2</t>", out2)) == 1,
      re.findall(r'<row r="(\d+)"', out2))

# chunk-boundary safety: a part larger than one CHUNK with rows straddling it
big_part = os.path.join(TD, "big_src.xml")
N = 60_000
with open(big_part, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:L{N + 3}"/><sheetData>')
    for r in range(1, N + 4):
        fv = 46100 + (r % 200)          # 46100..46299 ; >= 46235 -> 65 of every 200
        f.write(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>Client {r} '
                f'{"x" * 40}</t></is></c><c r="F{r}"><v>{fv}</v></c>'
                f'<c r="L{r}"><v>{r}</v></c></row>')
    f.write(f'</sheetData><autoFilter ref="A3:L{N + 3}"/></worksheet>')
size_mb = os.path.getsize(big_part) / 1048576
expected_drop = sum(1 for r in range(4, N + 4) if 46100 + (r % 200) >= 46235)
dropped = XlsxSurgeon._stream_trim_rows(big_part, os.path.join(TD, "big_out.xml"),
                                        "F", 46235.0, 4)
big_out = open(os.path.join(TD, "big_out.xml"), encoding="utf-8").read()
check(f"T9 multi-chunk stream ({size_mb:.1f} MB) drops exactly the matching rows",
      dropped == expected_drop, (dropped, expected_drop))
check("T10 multi-chunk stream keeps every non-matching row and the suffix",
      len(re.findall(r"<row r=", big_out)) == N + 3 - expected_drop
      and big_out.rstrip().endswith("</worksheet>"),
      len(re.findall(r"<row r=", big_out)))

# the real path on a part bigger than the 2MB tail window: trim + append via
# _transform_sheet must also re-pin <dimension> (lives in the head chunk)
big_dst = os.path.join(TD, "big_ta.xml")
s._transform_sheet(big_part, big_dst, {}, [(new_rows, None)], trim_groups=[("F", 46235.0, 4)])
with open(big_dst, encoding="utf-8") as f:
    head = f.read(2048)
    f.seek(max(0, os.path.getsize(big_dst) - 4096))
    tail = f.read()
last_kept = max(r for r in range(1, N + 4) if r < 4 or 46100 + (r % 200) < 46235)
final = last_kept + len(new_rows)
check("T11 large part: trim+append re-pins dimension and autoFilter to last kept + appended",
      f'<dimension ref="A1:L{final}"/>' in head and f'<autoFilter ref="A3:L{final}"/>' in tail,
      (re.findall(r'<dimension ref="[^"]*"', head), re.findall(r'<autoFilter ref="[^"]*"', tail), final))
check("T12 large part: trimmed rows counted",
      getattr(s, "_last_trimmed", None) == expected_drop, getattr(s, "_last_trimmed", None))


# ---------------------------------------------------------------- build_ops
spec = {"asofDate": "2026-08-31", "fromDate": "2026-08-01", "toDate": "2026-08-31",
        "ar": {"target": "AR_08.31", "anchor": "A7", "formulaCols": "Y:AH"},
        "sales": {"target": "New Sales report", "anchor": "A7", "formulaCols": "Z:AH"},
        "raw": {"target": "Sales report Raw", "anchor": "A4", "formulaCols": "Z:AA"}}
data = {"arRows": [[None] * 24], "salesRows": [[None] * 25]}
try:
    ops, report = build_ops(data, spec, prior_ar_tab="AR_07.31")
    kinds = [(o["op"], o.get("sheet")) for o in ops]
    ti = kinds.index(("trim_rows", "Sales report Raw"))
    ai = kinds.index(("append_rows", "Sales report Raw"))
    trim = ops[ti]
    check("B1 build_ops emits trim_rows on the raw tab BEFORE its append", ti < ai, kinds)
    check("B2 trim keyed on column F at the period-start serial, from the data anchor row",
          trim["col"] == "F" and trim["minValue"] == 46235 and trim["firstRow"] == 4, trim)
    check("B3 report records the trim serial", report.get("rawTrimFromSerial") == 46235, report)
except Exception as e:  # noqa: BLE001
    check("B1 build_ops emits trim_rows on the raw tab BEFORE its append", False, repr(e))

print("\n" + ("ALL PASS" if not fails else f"{len(fails)} FAIL: {fails}"))
sys.exit(1 if fails else 0)
