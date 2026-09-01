"""Mike/Preet 2026-09-01 review: blank AR exclusion, Harmony date-
conditional promo rates (adopted from Mike's own AR-tab edit, minus his
Excel session's external-workbook '[1]' rebinding and the accidental Dayna
comparison flip), and the prior-layout header labels."""
import sys

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
import netsuite_extract as ne
import commission_job as cj
import report_data as rd

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:300]}"))
    if not cond:
        fails.append(label)


PROMO_RAW = [
    {"sku": 98968, "pre": 0.15, "post": 0.10, "cutoffDate": "2026-08-01"},
    {"sku": 98900, "pre": 0.15, "post": 0.10, "cutoffDate": "2026-08-01"},
    {"sku": 99018, "pre": 0.12, "post": 0.10, "cutoffDate": "2026-08-01"},
    {"sku": 99017, "pre": 0.15, "post": 0.10, "cutoffDate": "2026-08-01"},
    {"sku": 99164, "pre": 0.10, "post": 0.10, "cutoffDate": "2026-08-01"},
]
promo = ne.normalize_promo(PROMO_RAW)
check("N1: cutoff date -> serial 46235 (Aug 1 2026)",
      all(p[3] == 46235 for p in promo), promo)

# ── template injection ──
tpl = ne.FORMULA_TEMPLATES["ar"]["AA"]
out = ne.inject_promo(tpl, promo)
check("I1: four branches injected in config order (99164 pre==post skipped)",
      out.count("IF(AND(I{r}=") == 4
      and out.index("98968") < out.index("98900") < out.index("99018")
      < out.index("99017") and "99164" not in out, out)
check("I2: branch shape matches Mike's edit",
      "IF(AND(I{r}=98968,F{r}<46235),0.15," in out
      and "IF(AND(I{r}=99018,F{r}<46235),0.12," in out, out)
check("I3: LOCAL SKU-tab reference kept (no '[1]' external rebinding)",
      "'Commission Rate by SKUs'!B:E" in out and "[1]" not in out, out)
check("I4: Dayna gate keeps the ORIGINAL direction (F>$AD$1)",
      "F{r}>$AD$1" in out and "F{r}<$AD$1" not in out, out)
check("I5: no promo -> template unchanged", ne.inject_promo(tpl, []) == tpl)

cells = ne.formula_cells("ar", 7, 2, prior_ar_tab="AR_07.31", promo=promo)
check("F1: generated AA7 carries the branches with real row numbers",
      "IF(AND(I7=98968,F7<46235),0.15," in cells["AA7"], cells["AA7"])
pcells = ne.prior_formula_cells(7, 1, promo=promo)
check("F2: prior-layout AD gets the same branches",
      "IF(AND(I7=98900,F7<46235),0.15," in pcells["AD7"], pcells["AD7"])
check("F3: prior AF/AG untouched by injection",
      pcells["AF7"].startswith("=IFERROR(IF(OR(AC7")
      and pcells["AG7"] == '=+CONCATENATE(H7," - ",I7)')

# ── shadow parity ──
CONSTS = {"dayna": "Dayna Stambeck", "ad1": 45703, "ai1": 46053,
          "aj4": 45762,
          "promo": {str(s): (p, c) for s, p, po, c in promo if p != po}}
SKU = {"98968": 0.10, "98900": 0.10, "99018": 0.10, "99017": 0.10,
       "99164": 0.10, "95302": 0.075}
def rate(item, f):
    return rd.shadow_rate("Tiffany McDaniel", "Acme Co", item, 726770, f,
                          SKU, set(SKU), CONSTS)
check("S1: pre-cutoff shipment -> promo rate", rate(98968, 46234) == 0.15)
check("S2: cutoff day itself -> post rate (strict <)", rate(98968, 46235) == 0.10)
check("S3: bundle 99018 pre-cutoff -> 0.12", rate(99018, 46200) == 0.12)
check("S4: 99164 has no branch -> tab rate both sides",
      rate(99164, 46200) == 0.10 and rate(99164, 46240) == 0.10)
check("S5: non-promo SKU unaffected", rate(95302, 46200) == 0.075)
check("S6: no promo in consts -> classic behavior",
      rd.shadow_rate("Tiffany McDaniel", "Acme Co", 98968, 1, 46200,
                     SKU, set(), {**CONSTS, "promo": {}}) == 0.10)

# ── blank exclusion: AR drops blank+None, sales keeps blank ──
def raw_row(ent):
    return {"c01": "815247 Cust", "c02": "6/15/2026", "c03": "Invoice",
            "c04": "726770", "c05": "95302", "c06": "-2", "c07": "-100",
            "c08": None, "c09": "Some Partner", "c10": "Open",
            "c11": "7/15/2026", "c12": None, "c13": "-100",
            "c14": "9", "c15": ent}
def cust(primary):
    return {"company": "Acme Co", "category": "Romane",
            "store_type": "Western", "type": "Retail", "first_sale": 44000,
            "partner": primary, "commission_pct": "10"}
CUST = {"1": cust("None"), "2": cust(None), "3": cust("Bob Logue"),
        "4": cust("  ")}
raw = [raw_row(e) for e in CUST]
ar, ar_s = ne.build_ar_rows(raw, CUST, {}, {}, {}, {}, sign_flip=True)
check("B1: AR keeps only the named partner",
      len(ar) == 1 and ar[0][20] == "Bob Logue"
      and ar_s["skipped_primary_none"] == 3, (len(ar), ar_s))
sales, s_s = ne.build_sales_rows(raw, CUST, {}, {}, 46265, sign_flip=True)
check("B2: sales keeps blanks, drops only 'None'",
      len(sales) == 3 and s_s["skipped_primary_none"] == 1, (len(sales), s_s))

# ── prior-layout header labels ──
ops, rep = cj.build_prior_ops(
    [[None] * 24 for _ in range(3)],
    {"target": "AR_07.31", "anchor": "A7", "promoRates": PROMO_RAW,
     "headerCells": {"AB4": 9}},
    46234)  # asof = Jul 31 2026
hc = next(o for o in ops if o["op"] == "set_cells")["cells"]
check("H1: labels match Mike's corrections",
      hc["AA6"] == "Deposit Month" and hc["AD6"] == "Commission Rate"
      and hc["AG6"] == "Concatenation & Item".replace("&", "No. &"),
      {k: hc.get(k) for k in ("AA6", "AD6", "AG6")})
check("H2: month-dynamic labels (July prior, August receipts)",
      hc["AE6"] == "July'26 Unearned Commission"
      and hc["AF6"] == "Commission earned on August'26 receipts",
      (hc.get("AE6"), hc.get("AF6")))
check("H3: 'Delete Later' column blanked", hc["AH6"] == " ")
check("H4: n8n headerCells still override (AB4)", hc["AB4"] == 9)
check("H5: prior AD rows carry promo branches",
      "IF(AND(I7=98968,F7<46235),0.15," in hc["AD7"], hc.get("AD7"))
# December rollover: prior asof Dec 31 2026 = 46387 -> receipts January'27
ops2, _ = cj.build_prior_ops([[None] * 24], {"target": "X", "anchor": "A7"},
                             46387)
hc2 = next(o for o in ops2 if o["op"] == "set_cells")["cells"]
check("H6: year rollover (Dec prior -> January'27 receipts)",
      hc2["AE6"] == "December'26 Unearned Commission"
      and hc2["AF6"] == "Commission earned on January'27 receipts",
      (hc2.get("AE6"), hc2.get("AF6")))

print()
print("ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
