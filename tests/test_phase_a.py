"""FOLLOW-UP 4 Part 1 tests: readRanges, headerGuard, newItems, inspect."""
import json, os, re, sys, zipfile, tempfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
os.environ.setdefault("NS_MCP_URL", "https://example.invalid/mcp")
import service
import xlsx_read as xr
import netsuite_extract as nx

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:280]}"))
    if not cond:
        fails.append(label)


TD = tempfile.mkdtemp(prefix="phase_a_")

AR_HDRS = service._AR_HDR
SALES_HDRS = [(h[0] if isinstance(h, tuple) else h) for h in service._SALES_HDR]

SHEETS = ["AR_06.30", "Dashboard", "New Sales report", "Sales report Raw",
          "Commission Rate by SKUs", "Data"]
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
      + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 7))
      + '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<sheets>' + "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(SHEETS, 1))
      + '</sheets><calcPr calcId="1"/></workbook>')
WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, 7))
           + '</Relationships>')

# shared strings: header names for the SALES tab live in sharedStrings (like
# the real Excel-authored file); AR headers are inline strings
SST_ITEMS = SALES_HDRS + ["SKU-A", "desc A", "SKU-B", "desc B"]
SST = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(SST_ITEMS)}" uniqueCount="{len(SST_ITEMS)}">'
       + "".join(f'<si><t>{h}</t></si>' for h in SST_ITEMS) + '</sst>')


def inline(ref, s):
    return f'<c r="{ref}" t="inlineStr"><is><t>{s}</t></is></c>'


def shared(ref, idx):
    return f'<c r="{ref}" t="s"><v>{idx}</v></c>'


def num(ref, v):
    return f'<c r="{ref}"><v>{v}</v></c>'


def frm(ref, f, v=None):
    return f'<c r="{ref}"><f>{f}</f>' + (f'<v>{v}</v>' if v is not None else '') + '</c>'


def ws(body, dim="A1:AH60"):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="{dim}"/><sheetData>{body}</sheetData></worksheet>')


ar_hdr_cells = "".join(inline(f"{xr.col_letter(i+1)}6", h) for i, h in enumerate(AR_HDRS))
ar_body = f'<row r="6">{ar_hdr_cells}</row><row r="7">{inline("A7", "data")}</row>'
# Dashboard W6:W24 with a gap at W10; C2 label
dash_cells = "".join(num(f"W{r}", r * 100.5) for r in range(6, 25) if r != 10)
dash_body = (f'<row r="2">{inline("C2", "June/26")}</row>'
             + "".join(f'<row r="{r}">{num(f"W{r}", r*100.5)}</row>' for r in range(6, 25) if r != 10))
sales_hdr_cells = "".join(shared(f"{xr.col_letter(i+1)}6", i) for i in range(len(SALES_HDRS)))
sales_body = f'<row r="6">{sales_hdr_cells}</row>'
raw_hdr_cells = "".join(shared(f"{xr.col_letter(i+1)}3", i) for i in range(len(SALES_HDRS)))
raw_body = f'<row r="3">{raw_hdr_cells}</row>'
sku_body = ('<row r="1">' + inline("B1", "Item") + inline("E1", "Rate") + '</row>'
            '<row r="2">' + shared("B2", len(SALES_HDRS)) + shared("C2", len(SALES_HDRS) + 1) + num("E2", 0.075) + '</row>'
            '<row r="3">' + num("B3", 98317) + shared("C3", len(SALES_HDRS) + 3) + num("E3", 0.1) + '</row>')
data_body = ('<row r="1">' + inline("A1", "Source") + inline("B1", "Amt") + '</row>'
             '<row r="2">' + inline("A2", "prior AR") + frm("B2", "SUM(1,2)", "3") + '</row>')

src = os.path.join(TD, "src.xlsx")
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/sharedStrings.xml", SST)
    for i, body in enumerate([ar_body, dash_body, sales_body, raw_body,
                              sku_body, data_body], 1):
        z.writestr(f"xl/worksheets/sheet{i}.xml", ws(body))

# ── headerGuard ──
guards = [{"sheet": "AR_07.31", "row": 6, "expect": "A:X"},
          {"sheet": "New Sales report", "row": 6, "expect": "A:Y"},
          {"sheet": "Sales report Raw", "row": 3, "expect": "A:Y"}]
ops = [{"op": "duplicate_sheet", "source": "AR_06.30", "name": "AR_07.31"}]
rep = service._run_header_guard(src, guards, ops)
check("G1: all three guards pass (dup source + sharedString headers)",
      len(rep) == 3 and all(r["status"] == "ok" for r in rep), rep)
check("G2: AR guard read from the dup source",
      rep[0]["readFrom"] == "AR_06.30", rep[0])

# shifted header must fail naming the column
shift_src = os.path.join(TD, "shift.xlsx")
bad_hdrs = list(AR_HDRS)
bad_hdrs.insert(16, bad_hdrs.pop(17))   # swap Q/R -> mismatch at Q
bad_cells = "".join(inline(f"{xr.col_letter(i+1)}6", h) for i, h in enumerate(bad_hdrs))
with zipfile.ZipFile(shift_src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/sharedStrings.xml", SST)
    for i, body in enumerate([f'<row r="6">{bad_cells}</row>', dash_body,
                              sales_body, raw_body, sku_body, data_body], 1):
        z.writestr(f"xl/worksheets/sheet{i}.xml", ws(body))
try:
    service._run_header_guard(shift_src, guards[:1], ops)
    check("G3: shifted layout raises", False)
except ValueError as e:
    check("G3: shifted layout raises naming sheet+column",
          "AR_07.31" in str(e) and "column Q" in str(e), str(e))

# ── readRanges ──
j = {}
service._run_read_ranges(src, [
    {"sheet": "Dashboard", "range": "W6:W24", "as": "endingBalances"},
    {"sheet": "Commission Rate by SKUs", "range": "B:E", "as": "skuRates"},
    {"sheet": "Nope", "range": "A1:A2", "as": "missing"},
], j)
eb = j["readRanges"]["endingBalances"]
check("R1: endingBalances = 19 rows incl. None gap",
      len(eb) == 19 and eb[0] == [6 * 100.5] and eb[4] == [None], eb[:6])
sk = j["readRanges"]["skuRates"]
check("R2: skuRates rows resolve shared strings + numbers",
      ["SKU-A", "desc A", None, 0.075] in sk and [98317, "desc B", None, 0.1] in sk, sk)
check("R3: missing sheet reported as error, not fatal",
      "missing" in j.get("readRangesErrors", {}), j.get("readRangesErrors"))

# ── inspect ──
info = service._run_inspect(src, {"rows": [{"sheet": "Data", "rows": "1:3"},
                                           {"sheet": "Dashboard", "rows": "1:5"}]})
check("I1: sheetNames + parts with dimensions",
      "Data" in info["sheetNames"] and info["parts"]["Data"]["dimension"] == "A1:AH60",
      info["parts"].get("Data"))
check("I2: row dump has values AND formulas",
      info["rows"]["Data"]["2"]["B"].get("f") == "=SUM(1,2)"
      and info["rows"]["Data"]["2"]["B"].get("v") == 3, info["rows"]["Data"])
check("I3: no pivots reported on pivot-free file", info.get("pivots") == [], info.get("pivots"))

# ── inspect: synthetic pivot parsing ──
piv_src = os.path.join(TD, "piv.xlsx")
PT = ('<pivotTableDefinition xmlns="x" name="SummaryPivot" cacheId="7">'
      '<rowFields count="2"><field x="0"/><field x="2"/></rowFields>'
      '<colFields count="1"><field x="-2"/></colFields>'
      '<dataFields count="1"><dataField name="Sum of Amount" fld="1"/></dataFields>'
      '</pivotTableDefinition>')
PCD = ('<pivotCacheDefinition xmlns="x"><cacheSource type="worksheet">'
       '<worksheetSource ref="A1:C100" sheet="Data"/></cacheSource>'
       '<cacheFields count="3"><cacheField name="Company"/><cacheField name="Amount"/>'
       '<cacheField name="Rate"/></cacheFields></pivotCacheDefinition>')
PT_RELS = ('<Relationships xmlns="r"><Relationship Id="rId1" Type="t" '
           'Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>')
with zipfile.ZipFile(piv_src, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/sharedStrings.xml", SST)
    for i, body in enumerate([ar_body, dash_body, sales_body, raw_body,
                              sku_body, data_body], 1):
        z.writestr(f"xl/worksheets/sheet{i}.xml", ws(body))
    z.writestr("xl/pivotTables/pivotTable1.xml", PT)
    z.writestr("xl/pivotTables/_rels/pivotTable1.xml.rels", PT_RELS)
    z.writestr("xl/pivotCache/pivotCacheDefinition1.xml", PCD)
pinfo = service._run_inspect(piv_src, {"rows": []})
pv = pinfo["pivots"][0]
check("I4: pivot parsed — name/cache/source/fields",
      pv["name"] == "SummaryPivot"
      and pv["cacheSource"] == {"sheet": "Data", "ref": "A1:C100"}
      and pv["rowFields"] == ["Company", "Rate"]
      and pv["colFields"] == ["(values)"]
      and pv["dataFields"] == ["Sum of Amount"], pv)

# ── newItems: mocked MCP two-step + fallback ──
class MockMcp:
    def __init__(self, fail_itemship=False):
        self.fail = fail_itemship
    def rows(self, q, label=""):
        if "'ItemShip'" in q and self.fail:
            raise nx.McpError("Record 'ItemShip' was not found")
        if "SELECT DISTINCT" in q:
            return [{"c01": "98317"}, {"c01": "99416"}]
        if "MIN(t.trandate)" in q:
            return [{"c01": "99416", "c02": "7/20/2026"}]
        if "i.description" in q:
            return [{"c01": "99416", "c02": "New  Thing"}]
        return []

items, basis = nx.fetch_new_items(MockMcp(), "2026-07-01", "2026-07-31", log=lambda *a: None)
check("N1: first-fulfilled filter + ISO date",
      items == [("99416", "2026-07-20")] and basis == "item_fulfillment", (items, basis))
items2, basis2 = nx.fetch_new_items(MockMcp(fail_itemship=True), "2026-07-01", "2026-07-31", log=lambda *a: None)
check("N2: ItemShip-blocked falls back with labelled basis",
      basis2 == "first_sale_fallback" and items2 == [("99416", "2026-07-20")], (items2, basis2))

# ── Job model accepts the v10 payload shapes ──
jm = service.Job(downloadUrl="d", uploadUrl="u",
                 readRanges=[{"sheet": "Dashboard", "range": "W6:W24", "as": "endingBalances"}],
                 headerGuard=[{"sheet": "X", "row": 6, "expect": "A:X"}],
                 inspect={"rows": []})
check("M1: model accepts readRanges/headerGuard/inspect", jm.readRanges is not None)
jm2 = service.Job(downloadUrl="d", uploadUrl="u", headerGuard=None, readRanges=None)
check("M2: nulls accepted (v10 sends headerGuard: null when disabled)", jm2.headerGuard is None)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
