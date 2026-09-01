"""Mike/Preet review 2026-08-25, items 1 and 4:

1. Rows whose PRIMARY partner is the literal partner record "None"
   (Target/Walmart-style house accounts) are excluded from the AR and
   sales extracts; BLANK primary partner stays in for now.
2. The shadow prior-balance lookup keys must be Excel-faithful: CONCAT
   does not trim, XLOOKUP's exact match is case-insensitive.
"""
import sys

sys.path.insert(0, r"C:\Users\DanielMirochnik\Marketing MCP Servers\commission-surgeon")
from netsuite_extract import build_ar_rows, build_sales_rows
import report_data as rd

fails = []


def check(label, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + label + ("" if cond else f"  {str(detail)[:260]}"))
    if not cond:
        fails.append(label)


def raw_row(ent):
    return {"c01": "815247 Cust", "c02": "6/15/2026", "c03": "Invoice",
            "c04": "726770", "c05": "95302", "c06": "-2", "c07": "-100",
            "c08": None, "c09": "Some Partner", "c10": "Open",
            "c11": "7/15/2026", "c12": None, "c13": "-100",
            "c14": "9", "c15": ent}


def cust_entry(primary):
    return {"company": "Acme Co", "category": "Romane",
            "store_type": "Western", "type": "Retail", "first_sale": 44000,
            "partner": primary, "commission_pct": "10"}


CUST = {"1": cust_entry("None"),            # literal partner record "None"
        "2": cust_entry(None),              # BLANK -> kept, per Mike
        "3": cust_entry("Tiffany McDaniel"),
        "4": cust_entry(" none ")}          # whitespace/case variant

raw = [raw_row("1"), raw_row("2"), raw_row("3"), raw_row("4")]

ar, ar_stats = build_ar_rows(raw, CUST, {}, {}, {}, {}, sign_flip=True)
check("A1: AR keeps blank + named, drops both None variants",
      len(ar) == 2 and {r[20] for r in ar} == {None, "Tiffany McDaniel"},
      [(r[20]) for r in ar])
check("A2: AR stats count the exclusions",
      ar_stats["skipped_primary_none"] == 2, ar_stats)

sales, s_stats = build_sales_rows(raw, CUST, {}, {}, 46234, sign_flip=True)
check("S1: sales same rule", len(sales) == 2
      and s_stats["skipped_primary_none"] == 2, s_stats)

# ── Excel-faithful lookup keys ──
CONSTS = {"dayna": "Dayna Stambeck", "ad1": 45703, "ai1": 46053,
          "aj4": 45762}


def ar_row(company, doc="726770", item="95302", bal=100.0, gross=100.0,
           closed=None, partner="Tiffany McDaniel"):
    r = [None] * 24
    r[0] = "815247 Acme"; r[3] = 44000.0; r[7] = doc; r[8] = item
    r[11] = bal; r[19] = closed; r[20] = partner; r[21] = gross
    r[23] = company
    return r


SKU = {"95302": 0.1}

# prior holds "ACME Co " (trailing space, caps); current month re-extracts
# the same doc as "acme Co " — Excel's CONCAT keeps the space on BOTH
# sides and XLOOKUP matches case-insensitively, so this must HIT.
prior = [ar_row("ACME Co ", bal=100.0)]
cur = [ar_row("acme Co ", bal=30.0, gross=100.0)]
out, pp = rd.shadow_compiled(prior, cur, [], 46234, SKU, set(), CONSTS)
check("K1: case-insensitive hit -> partial = prior 100 - current 30 = 70",
      len(pp) == 1 and abs(pp[0][2] - 70.0) < 0.005, pp)

# different INTERNAL spacing is a genuinely different CONCAT string — Excel
# misses, so we must miss too (fallback AL = AK gross)
prior = [ar_row("Other  LLC", bal=100.0)]
cur = [ar_row("Other LLC", bal=30.0, gross=80.0)]
out, pp = rd.shadow_compiled(prior, cur, [], 46234, SKU, set(), CONSTS)
check("K2: different spacing misses like Excel -> AL falls back to gross",
      len(pp) == 1 and abs(pp[0][2] - 50.0) < 0.005, pp)

# trailing-space asymmetry: one side padded, the other not — Excel keeps
# the space, so this MISSES (the old .strip() version wrongly matched it)
prior = [ar_row("Padded Co ", bal=100.0)]
cur = [ar_row("Padded Co", bal=30.0, gross=80.0)]
out, pp = rd.shadow_compiled(prior, cur, [], 46234, SKU, set(), CONSTS)
check("K3: trailing-space asymmetry misses like Excel",
      len(pp) == 1 and abs(pp[0][2] - 50.0) < 0.005, pp)

# int/float/str doc renderings still key identically
prior = [ar_row("Same Co", doc=726770)]
cur = [ar_row("Same Co", doc=726770.0, bal=40.0, gross=100.0)]
out, pp = rd.shadow_compiled(prior, cur, [], 46234, SKU, set(), CONSTS)
check("K4: 726770 vs 726770.0 keys identically -> partial 60",
      len(pp) == 1 and abs(pp[0][2] - 60.0) < 0.005, pp)

print()
print("ALL PASS" if not fails else f"{len(fails)} FAILURES: {fails}")
sys.exit(1 if fails else 0)
