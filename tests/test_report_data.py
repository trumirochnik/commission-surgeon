"""report_data (the reporting half) unit tests — Data-tab builder,
Compiled combo rows, rate engine, statement files."""
import io, json, re, sys, zipfile

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
import report_data as rd

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:260]}"))
    if not cond:
        fails.append(label)


ASOF = 46234   # July 31 2026


def ar_row(company="Acme Co", partner="Tiffany McDaniel", item="95302",
           doc="726770", closed=None, first_sale=44000.0, gross=100.0,
           bal=100.0):
    r = [None] * 24
    r[0] = "815247 Acme"; r[3] = first_sale; r[7] = doc; r[8] = item
    r[11] = bal; r[19] = closed; r[20] = partner; r[21] = gross
    r[23] = company
    return r


def sales_row(company="Acme Co", partner="Tiffany McDaniel", item="95302"):
    r = [None] * 25
    r[0] = "815247 Acme"; r[7] = "9001"; r[8] = item
    r[11] = 50.0; r[20] = partner; r[21] = 50.0
    r[24] = company
    return r


prior = [ar_row(closed=46210.0),            # closed in July  -> AB set
         ar_row(closed=46240.0),            # closed in Aug   -> AB REMOVED
         ar_row(closed=None)]               # open            -> AB empty
cur = [ar_row(), ar_row()]
sales = [sales_row()]

b = rd.build_data_rows(prior, cur, sales, ASOF, "AR_06.30", "AR_07.31",
                       "New Sales report", "Commission earned on JUL'26")
rows, cells = b["pasteRows"], b["cells"]

check("row count = blocks + spacers",
      len(rows) == 3 + rd.SPACER + 2 + rd.SPACER + 1, len(rows))
check("block boundaries", b["blocks"] == {"prior": (11, 13),
      "current": (16, 17), "sales": (20, 20)}, b["blocks"])
check("period labels exact",
      rows[0][1] == "Prior Month" and rows[5][1] == "Current Month"
      and rows[9][1] == "New Sales", [r[1] for r in rows])
check("spacer rows are empty", rows[3] == [None] * 26 and rows[8] == [None] * 26)
check("C..Z carry the source's 24 cols", rows[0][2] == "815247 Acme"
      and rows[0][15 + 2] is None and rows[0][2 + 23] == "Acme Co")
check("sales row: Z=company, Y empty",
      rows[9][25] == "Acme Co" and rows[9][24] is None, rows[9])

check("prior AB: July close kept, Aug close removed, open empty",
      cells.get("AB11") == 46210.0 and "AB12" not in cells
      and "AB13" not in cells and cells.get("AD12") == 0,
      {k: cells.get(k) for k in ("AB11", "AB12", "AD12")})
check("prior AD month number", cells.get("AD11") == 7)
check("prior AE/AF/AG reference the refreshed tab",
      cells["AE11"] == "='AR_06.30'!AD7" and cells["AG13"] == "='AR_06.30'!AF9")
check("current AA is the SOP formula", cells["AA16"] == "=X16-N16")
check("current AJ pinned to prior block extent",
      "$AI$11:$AI$13" in cells["AJ16"] and "$N$11:$N$13" in cells["AJ16"],
      cells["AJ16"])
check("current AK original gross value + X = adjusted formula",
      cells["AK16"] == 100.0 and cells["X16"] == "=AL16")
check("sales AB..AG reference the sales tab",
      cells["AB20"] == "='New Sales report'!AB7"
      and cells["AC20"] == "='New Sales report'!X7"
      and cells["AG20"] == "='New Sales report'!AF7")
check("keys everywhere", cells["AI11"].startswith("=_xlfn.CONCAT(Z11")
      and cells["AM20"] == '=_xlfn.CONCAT(J20," - ",K20)')
check("header cells", b["headerCells"] == {"AD4": ASOF, "AB8": ASOF + 1,
      "AG10": "Commission earned on JUL'26"})

# ── combos + rate engine ──
sku = {"95302": 0.075}
kev = {"726770": 0.05}
combos, undet = rd.distinct_combos(prior, cur, sales, sku, kev)
check("combo detected with SKU rate",
      ("Acme Co", "Tiffany McDaniel", 0.075) in combos, combos)
combos2, _ = rd.distinct_combos([ar_row(partner="Kevin Hanks")], [], [], sku, kev)
check("Kevin rate from contracted lookup",
      ("Acme Co", "Kevin Hanks", 0.05) in combos2, combos2)
combos3, _ = rd.distinct_combos([ar_row(partner="None")], [], [], {}, {})
check("None partner excluded (rate 0)", not combos3, combos3)
combos4, _ = rd.distinct_combos([ar_row(item="UNKNOWN")], [], [], sku, kev)
check("unknown SKU -> default 10%",
      ("Acme Co", "Tiffany McDaniel", 0.1) in combos4, combos4)

cc = rd.compiled_combo_rows([("Acme Co", "Tiffany McDaniel", 0.075)],
                            2183, 46000, "JUL'26")
check("combo row cells complete",
      cc["C2183"] == "Acme Co" and cc["F2183"] == 0.075
      and "Data!$N$11:$N$46000" in cc["G2183"]
      and '">0"' in cc["I2183"] and cc["O2183"] == "=M2183-N2183",
      {k: str(v)[:60] for k, v in cc.items() if "2183" in k})

# ── statements ──
compiled = [
    {"company": "Acme Co", "partner": "Tiffany McDaniel", "rate": 0.1,
     "prior": 100, "newSales": 50, "collections": 80, "partial": 5,
     "totalColl": 85, "earned": 8.5},
    {"company": "Beta LLC", "partner": "Kevin Hanks", "rate": 0.05,
     "prior": 10, "newSales": 0, "collections": 10, "partial": 0,
     "totalColl": 10, "earned": 0.5},
    {"company": "Beta LLC", "partner": "Kevin Hanks", "rate": 0.075,
     "prior": 20, "newSales": 0, "collections": 20, "partial": 0,
     "totalColl": 20, "earned": 1.5},
    {"company": "Idle Co", "partner": "Pam Winningham", "rate": 0.1,
     "prior": 0, "newSales": 0, "collections": 0, "partial": 0,
     "totalColl": 0, "earned": 0},
]
payment = {"Tiffany McDaniel": {"earned": 8.5, "fee": 100, "adj": 0, "net": -8.5}}
files = rd.build_statements(compiled, payment, "07.2026")
check("one file per active partner, idle partner skipped",
      sorted(files) == ["07.2026_Commission Statement_Kevin Hanks.xlsx",
                        "07.2026_Commission Statement_Tiffany McDaniel.xlsx"],
      sorted(files))
tz = zipfile.ZipFile(io.BytesIO(files["07.2026_Commission Statement_Kevin Hanks.xlsx"]))
check("statement is a valid xlsx", tz.testzip() is None
      and "xl/worksheets/sheet1.xml" in tz.namelist())
sx = tz.read("xl/worksheets/sheet1.xml").decode()
check("Kevin aggregated by client (rates summed, one Beta row)",
      sx.count("Beta LLC") == 1 and ">30</v>" in sx and ">2.0</v>" in sx,
      sx[:800])
tz2 = zipfile.ZipFile(io.BytesIO(files["07.2026_Commission Statement_Tiffany McDaniel.xlsx"]))
sx2 = tz2.read("xl/worksheets/sheet1.xml").decode()
check("Tiffany statement carries payment block",
      "Net Payment" in sx2 and "Technology Fee" in sx2, sx2[-400:])

print()
if fails:
    print(f"{len(fails)} FAILURES: {fails}")
    sys.exit(1)
print("ALL PASS")
