"""
Nova Leisure Group — synthetic operational data generator.
Produces three CSVs that mimic a real Xero / POS export, with
intentional dirt (miscodes, missing tags, blanks, case noise).
"""
import csv, random, datetime as dt, os

random.seed(42)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SITES = [
    ("S001", "Nova London Bridge",   "London",     1.00),
    ("S002", "Nova Manchester Arndale", "Manchester", 0.78),
    ("S003", "Nova Birmingham Bullring","Birmingham", 0.82),
    ("S004", "Nova Bristol Cabot",   "Bristol",     0.62),
    ("S005", "Nova Edinburgh Princes","Edinburgh",  0.70),
]

DEPTS = ["FOOD_BEV", "RETAIL", "ENTERTAINMENT", "MEMBERSHIP", "EVENTS"]

# ---------- 1. FY25 Approved Budget --------------------------------
BUDGET_LINES = [
    # (account_code, description, category, jan_value, seasonality_pattern)
    ("4000", "Food & Beverage Revenue",  "Revenue",       620000,  "leisure"),
    ("4010", "Retail Merchandise Revenue","Revenue",      210000,  "retail"),
    ("4020", "Entertainment Revenue",    "Revenue",       380000,  "leisure"),
    ("4030", "Membership Revenue",       "Revenue",       145000,  "flat"),
    ("4040", "Events & Hire Revenue",    "Revenue",        92000,  "events"),
    ("5000", "Cost of Goods Sold - F&B", "COGS",         -198000,  "leisure"),
    ("5010", "Cost of Goods Sold - Retail","COGS",       -115000,  "retail"),
    ("6000", "Wages & Salaries",         "OPEX_Wages",   -385000,  "flat"),
    ("6010", "Employer NI & Pension",    "OPEX_Wages",    -52000,  "flat"),
    ("6100", "Property Rent",            "OPEX_Fixed",    -140000,  "flat"),
    ("6110", "Business Rates",           "OPEX_Fixed",     -41000,  "flat"),
    ("6200", "Utilities - Electricity",  "OPEX_Variable",  -34000,  "energy"),
    ("6210", "Utilities - Gas/Water",    "OPEX_Variable",  -11000,  "energy"),
    ("6300", "Marketing & Promotions",   "OPEX_Variable",  -28000,  "events"),
    ("6400", "Repairs & Maintenance",    "OPEX_Variable",  -18000,  "flat"),
    ("6500", "IT & Software Licences",   "OPEX_Fixed",     -14000,  "flat"),
    ("6600", "Insurance",                "OPEX_Fixed",      -9000,  "flat"),
    ("6700", "Professional Fees",        "OPEX_Variable",   -7500,  "flat"),
    ("7000", "Depreciation",             "Non-Cash",       -42000,  "flat"),
    ("7100", "Bank Interest",            "Finance",         -8500,  "flat"),
]

def seasonality(pattern, month):
    # month is 1..12
    if pattern == "leisure":
        return [0.85,0.82,0.90,1.00,1.10,1.05,1.20,1.25,1.00,0.95,1.05,1.40][month-1]
    if pattern == "retail":
        return [0.70,0.75,0.85,0.95,1.00,1.00,1.05,1.05,0.95,1.05,1.20,1.65][month-1]
    if pattern == "events":
        return [0.60,0.70,0.85,1.00,1.10,1.20,1.15,0.95,1.05,1.10,1.20,1.60][month-1]
    if pattern == "energy":
        return [1.35,1.30,1.15,0.95,0.80,0.75,0.70,0.70,0.85,1.00,1.20,1.35][month-1]
    return 1.00

# Write FY25 Budget CSV
budget_path = os.path.join(ROOT, "FY25_Approved_Budget.csv")
with open(budget_path, "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Account_Code","Description","Category","Period","Budget_GBP"])
    for code, desc, cat, jan_val, patt in BUDGET_LINES:
        for m in range(1, 13):
            period = f"FY25-{m:02d}"
            val = round(jan_val * seasonality(patt, m))
            w.writerow([code, desc, cat, period, val])

# ---------- 2. YTD Actuals Ledger (dirty) --------------------------
# 12 months of high-volume transactions, with deliberate noise:
# - miscoded OPEX (account codes that don't exist in budget)
# - missing department tags
# - mixed-case descriptions, trailing spaces
# - duplicate rows, blank rows

actuals_path = os.path.join(ROOT, "YTD_Actuals_Ledger.csv")

# Macro hit: actuals run ~2-4% above budget on costs, mixed on revenue
def jitter(base, lo=-0.06, hi=0.04):
    return base * (1 + random.uniform(lo, hi))

rows = []
TXN_ID = 100000
for m in range(1, 13):
    period = f"FY25-{m:02d}"
    days_in_month = 31 if m in (1,3,5,7,8,10,12) else (28 if m == 2 else 30)
    for code, desc, cat, jan_val, patt in BUDGET_LINES:
        month_budget = jan_val * seasonality(patt, m)
        # split into multiple transactions per site
        for site_code, site_name, region, site_weight in SITES:
            # Skip some lines for some sites realistically
            if cat == "Finance" and site_code != "S001":
                continue
            site_total = month_budget * site_weight / sum(s[3] for s in SITES)
            # actualise with jitter (costs higher, revenue mixed)
            if cat == "Revenue":
                actual = jitter(site_total, -0.10, 0.06)
            elif cat in ("COGS", "OPEX_Wages", "OPEX_Variable", "OPEX_Fixed"):
                actual = jitter(site_total, -0.02, 0.08)  # costs over-run
            else:
                actual = jitter(site_total, -0.03, 0.03)
            # Split into 3-15 transactions
            n_txn = random.randint(3, 12) if abs(actual) > 5000 else random.randint(1,3)
            for i in range(n_txn):
                TXN_ID += 1
                amt = round(actual / n_txn, 2)
                txn_day = random.randint(1, days_in_month)
                txn_date = dt.date(2025, m, txn_day).isoformat()
                # Introduce dirt:
                d = desc
                if random.random() < 0.04: d = d.upper()
                if random.random() < 0.03: d = d + "  "  # trailing space
                if random.random() < 0.02: d = " " + d
                dept = random.choice(DEPTS) if cat in ("Revenue","COGS") else ""
                if random.random() < 0.06: dept = ""  # missing tag
                # Miscoded OPEX: 3% chance of wrong account code
                acc = code
                if cat.startswith("OPEX") and random.random() < 0.03:
                    acc = random.choice(["6999","9000","5099"])  # not in budget
                rows.append([TXN_ID, txn_date, period, acc, d, site_code, dept, amt])

# Add ~30 fully blank rows scattered through
for _ in range(30):
    pos = random.randint(0, len(rows))
    rows.insert(pos, ["","","","","","","",""])

# Add ~25 duplicate rows
for _ in range(25):
    src = random.choice([r for r in rows if r[0] != ""])
    rows.append(src.copy())

# Shuffle for realism
random.shuffle(rows)

with open(actuals_path, "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Txn_ID","Txn_Date","Period","Account_Code","Description","Site_Code","Department","Amount_GBP"])
    for r in rows:
        w.writerow(r)

# ---------- 3. Daily Site Footfall --------------------------------
footfall_path = os.path.join(ROOT, "Daily_Site_Footfall.csv")
start = dt.date(2025, 1, 1)
end   = dt.date(2026, 5, 25)  # ~17 months
days = (end - start).days + 1

# Footfall base per site (daily avg)
BASE = {"S001": 4200, "S002": 3100, "S003": 3400, "S004": 2200, "S005": 2600}

with open(footfall_path, "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Date","Site_Code","Site_Name","Region","Visitors","Avg_Spend_per_Head_GBP","Weather"])
    for i in range(days):
        d = start + dt.timedelta(days=i)
        for site_code, site_name, region, weight in SITES:
            # day-of-week effect (Fri/Sat higher)
            dow = d.weekday()  # 0=Mon
            dow_factor = [0.85, 0.85, 0.90, 0.95, 1.20, 1.45, 1.30][dow]
            # month seasonality
            month_factor = seasonality("leisure", d.month)
            # macro decline overlay (footfall declining YoY in 2026)
            if d.year == 2026:
                decline = 1 - 0.05 * (d.month / 12)  # progressive
            else:
                decline = 1.0
            # weather random
            weather_roll = random.random()
            if weather_roll < 0.15:
                weather = "Rain"; wfac = 0.82
            elif weather_roll < 0.30:
                weather = "Cold"; wfac = 0.92
            elif weather_roll < 0.85:
                weather = "Dry"; wfac = 1.00
            else:
                weather = "Sunny"; wfac = 1.12
            # holidays / events
            event_factor = 1.0
            if d.month == 12 and d.day in (24,26,27,28,29,30):
                event_factor = 1.6
            if d.month == 12 and d.day == 25:
                event_factor = 0.0  # closed Christmas day
            if d.month == 1 and d.day == 1:
                event_factor = 0.3
            visitors = int(BASE[site_code] * dow_factor * month_factor * decline * wfac * event_factor * random.uniform(0.92, 1.08))
            spend = round(random.uniform(11.0, 18.5) * (1.05 if dow >= 4 else 1.0), 2)
            w.writerow([d.isoformat(), site_code, site_name, region, visitors, spend, weather])

print(f"Wrote:\n  {budget_path}\n  {actuals_path}\n  {footfall_path}")
print(f"Actuals rows: {len(rows)}  Footfall rows: {days * len(SITES)}")
