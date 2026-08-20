"""Tests for the five data-defect fixes: sales W/X/Y tail, sales P, Q from
custentity12, AR W account, retarget_refs op."""
import os, re, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
import netsuite_extract as nx
from surgeon import XlsxSurgeon

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {detail}"))
    if not cond:
        fails.append(label)


# ── helpers ──
check("H1: month_label M/D/YYYY", nx.month_label("6/15/2026") == "Jun 2026")
check("H2: month_label ISO", nx.month_label("2022-10-03") == "Oct 2022")
check("H3: month_label None", nx.month_label(None) is None)

# ── row builders on synthetic raw rows ──
raw_row = {"c01": "875001 Store", "c02": "7/15/2026", "c03": "Invoice",
           "c04": "730001", "c05": "95302", "c06": "-4", "c07": "-100.5",
           "c08": "EMP204", "c09": "Tiffany McDaniel", "c10": "Invoice : Open",
           "c11": "8/15/2026", "c12": None, "c13": "-25.0", "c14": "29506822",
           "c15": "22373"}
cust = {"22373": {"type": "Wholesale", "category": "Romane",
                  "first_sale": nx.serial("2/9/2006"), "store_type": "Western",
                  "company": "Johnson's Men's Wear", "partner": "Tiffany McDaniel",
                  "commission_pct": "10"}}
items = {"95302": "Tru Western - Yellowstone"}
states = {"29506822": "TX"}
partners = {"Tiffany McDaniel": "Primary Rep"}
accounts = {"29506822": "11300 - Accounts Receivable - Trade"}
asof_serial = nx.serial("2026-07-31")

ar_rows, _ = nx.build_ar_rows([dict(raw_row)], cust, items, states, partners,
                              accounts, sign_flip=True)
r = ar_rows[0]
check("AR1: 24 values", len(r) == 24, len(r))
check("AR2: Q = commission pct label TEXT", r[16] == "10", repr(r[16]))
check("AR3: W = dashed AR account", r[22] == "11300 - Accounts Receivable - Trade", r[22])
check("AR4: P = partner role (unchanged)", r[15] == "Primary Rep", r[15])
check("AR5: sign flip still applied", r[10] == 4.0 and r[11] == 100.5, (r[10], r[11]))

sales_rows, _ = nx.build_sales_rows([dict(raw_row)], cust, items, states,
                                    asof_serial, sign_flip=True)
s = sales_rows[0]
check("S1: 25 values", len(s) == 25, len(s))
check("S2: P = customer's partner (person)", s[15] == "Tiffany McDaniel", s[15])
check("S3: Q = commission pct label TEXT", s[16] == "10", repr(s[16]))
check("S4: U unchanged (customer's partner)", s[20] == "Tiffany McDaniel", s[20])
check("S5: W = month label of row date", s[22] == "Jul 2026", s[22])
age = s[23]
check("S6: X = client age years ~20.47", isinstance(age, float) and 20.0 < age < 21.0, age)
check("S7: Y = company name", s[24] == "Johnson's Men's Wear", s[24])
check("S8: no tl.rate anywhere (dropped)", "-25.0" not in [str(v) for v in s], s)
check("S9: q_sales_lines no longer selects tl.rate",
      "tl.rate" not in nx.q_sales_lines(["1"]))
check("S10: q_customers selects custentity12 via DF",
      "BUILTIN.DF(cu.custentity12)" in nx.q_customers(["1"]))
check("S11: q_accounts is header-only, keyed",
      "transactionaccountingline" not in nx.q_accounts(["1"])
      and "BUILTIN.DF(t.account)" in nx.q_accounts(["1"]))

# ── retarget_refs op ──
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheets><sheet name="AR_06.30" sheetId="1" r:id="rId1"/></sheets>'
      '<calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '</Relationships>')
# June-style sheet: L2 is the load-bearing header formula the delivered file
# showed surviving ('AR_05.31'!V4 - 'AR_06.30'!L3 — prior-month total minus
# own L3). Data rows 7-8 carry May XLOOKUPs.
AR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
     '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
     '<dimension ref="A1:AH9"/><sheetData>'
     '<row r="2"><c r="L2"><f>\'AR_05.31\'!V4-\'AR_06.30\'!L3</f><v>19562.35</v></c></row>'
     '<row r="4"><c r="AD4"><f>\'AR_05.31\'!$B$1+0</f></c></row>'
     '<row r="6"><c r="A6" t="inlineStr"><is><t>Client</t></is></c></row>'
     '<row r="7"><c r="A7" t="inlineStr"><is><t>old</t></is></c>'
     '<c r="AD7"><f>+IFERROR(IF(F7&lt;$AD$4,_xlfn.XLOOKUP(AC7,\'AR_05.31\'!AG:AG,\'AR_05.31\'!L:L),IF(F7&gt;$AD$4,V7," ")),0)</f></c>'
     '<c r="AG7"><f>_xlfn.XLOOKUP(AC7,AR_05.31!AG:AG,AR_05.31!L:L)</f></c></row>'
     '<row r="8"><c r="A8" t="inlineStr"><is><t>old2</t></is></c></row>'
     '</sheetData></worksheet>')

WORK = tempfile.mkdtemp(prefix="retarget_")
src = os.path.join(WORK, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", AR)


def sheet_of(path, name):
    with zipfile.ZipFile(path) as z:
        wb = z.read("xl/workbook.xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
        rid = re.search(rf'name="{re.escape(name)}"[^>]*r:id="([^"]+)"', wb).group(1)
        tgt = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1)
        return z.read("xl/" + tgt).decode()


# R1: dup + paste + set + TWO-WAY retarget (streaming path). The month
# roll: May->June AND June->July, simultaneously. Generated data rows
# reference AR_06.30 (as the templates emit) and must NOT be shifted.
TWO_WAY = [{"from": "AR_05.31", "to": "AR_06.30"},
           {"from": "AR_06.30", "to": "AR_07.31"}]
out1 = os.path.join(WORK, "out1.xlsx")
sg = XlsxSurgeon(src, workdir=WORK)
sg.duplicate_sheet("AR_06.30", "AR_07.31")
sg.paste_columns("AR_07.31", "A7", [["july1"] + [None] * 23,
                                    ["july2"] + [None] * 23], True)
sg.set_cells("AR_07.31", {
    "AD7": "=_xlfn.XLOOKUP(AC7,'AR_06.30'!AC:AC,'AR_06.30'!L:L)",
    "AD8": "=_xlfn.XLOOKUP(AC8,'AR_06.30'!AC:AC,'AR_06.30'!L:L)"})
sg.retarget_refs("AR_07.31", TWO_WAY)
res = sg.apply(out1)
dup = sheet_of(out1, "AR_07.31")
check("R1: no AR_05.31 left anywhere in dup (incl. rows 1-6)",
      "AR_05.31" not in dup, dup[:500])
check("R2: L2 two-way shifted correctly (no cascade)",
      "'AR_06.30'!V4-'AR_07.31'!L3" in dup, dup[:500])
check("R2b: header AD4 shifted May->June", "'AR_06.30'!$B$1" in dup)
check("R2c: generated data formulas NOT shifted (still prior-month June)",
      "_xlfn.XLOOKUP(AC7,'AR_06.30'!AC:AC" in dup.replace("&apos;", "'"), dup[-900:])
rt = [r0 for r0 in res if r0["kind"] == "retarget_refs"]
check("R3: per-mapping counts reported",
      rt and rt[0]["perMapping"].get("AR_05.31", 0) >= 1
      and rt[0]["perMapping"].get("AR_06.30", 0) >= 1, rt)
check("R4: source sheet untouched", "AR_05.31" in sheet_of(out1, "AR_06.30"))

# R5: retarget on EXISTING sheet (in-memory path, no paste)
out2 = os.path.join(WORK, "out2.xlsx")
sg = XlsxSurgeon(src, workdir=WORK)
sg.retarget_refs("AR_06.30", [{"from": "AR_05.31", "to": "AR_04.30"}])
res2 = sg.apply(out2)
ex = sheet_of(out2, "AR_06.30")
check("R5: existing-sheet retarget applied", "AR_05.31" not in ex and "'AR_04.30'!AG:AG" in ex)
check("R6: unquoted form also fixed", "AR_04.30!AG:AG" in ex, ex[-500:])
rt2 = [r0 for r0 in res2 if r0["kind"] == "retarget_refs"]
check("R7: count = 6 (L2 + AD4 + AD7x2 quoted, AG7x2 unquoted)",
      rt2 and rt2[0]["replacements"] == 6, rt2)

# R8: any mapping with zero replacements FAILS the job — even when the
# other mapping in the same op hits
sg = XlsxSurgeon(src, workdir=WORK)
sg.retarget_refs("AR_06.30", [{"from": "AR_05.31", "to": "AR_04.30"},
                              {"from": "AR_01.31", "to": "AR_02.28"}])
try:
    sg.apply(os.path.join(WORK, "out3.xlsx"))
    check("R8: zero-count mapping raises", False)
except ValueError as e:
    check("R8: zero-count mapping raises",
          "AR_01.31" in str(e) and "0 replacements" in str(e), str(e))

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
