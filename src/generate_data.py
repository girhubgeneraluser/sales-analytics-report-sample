"""Generate SYNTHETIC raw data for the Northwind Home Goods sample project.

Sample project - synthetic data, not client work. Seeded, so every run is identical.

Writes to data/raw/:
  shopify_orders_export.csv  one row per order line item (Shopify-style), with deliberate mess
  ad_spend_daily.csv         daily spend per channel x campaign (Meta + Google)
  products.csv               product catalog with unit cost
  fx_rates_monthly.csv       monthly CAD->USD rates
"""
from __future__ import annotations

import csv
import datetime as dt
import math
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np

SEED = 20260401
ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
STORE_TZ = ZoneInfo("America/New_York")
START = dt.date(2025, 4, 1)
END = dt.date(2026, 9, 30)

rng = np.random.default_rng(SEED)

# --------------------------------------------------------------------------- catalog
CATALOG = [
    # sku, name, category, price_usd, cost_usd, popularity, refund_prob
    ("KIT-001", "Stoneware Dinner Plate Set (4)", "Kitchen", 64.0, 22.0, 1.4, 0.04),
    ("KIT-002", "Cast Iron Skillet 10in", "Kitchen", 48.0, 19.0, 1.2, 0.03),
    ("KIT-003", "Linen Tea Towels (3)", "Kitchen", 24.0, 6.5, 1.8, 0.03),
    ("KIT-004", "Olivewood Serving Board", "Kitchen", 58.0, 21.0, 0.9, 0.05),
    ("KIT-005", "Glass Storage Jars (6)", "Kitchen", 39.0, 14.0, 1.1, 0.09),
    ("KIT-006", "Ceramic Pour-Over Set", "Kitchen", 72.0, 30.0, 0.6, 0.06),
    ("KIT-007", "Copper Measuring Cups", "Kitchen", 34.0, 15.0, 0.4, 0.05),
    ("KIT-008", "Bamboo Utensil Set", "Kitchen", 22.0, 7.0, 0.9, 0.03),
    ("BED-001", "Linen Duvet Cover - Queen", "Bedding", 189.0, 68.0, 1.0, 0.19),
    ("BED-002", "Percale Sheet Set - Queen", "Bedding", 129.0, 41.0, 1.5, 0.06),
    ("BED-003", "Percale Sheet Set - King", "Bedding", 149.0, 48.0, 0.9, 0.06),
    ("BED-004", "Waffle Knit Blanket", "Bedding", 119.0, 39.0, 1.0, 0.05),
    ("BED-005", "Down-Alt Pillow (2)", "Bedding", 79.0, 26.0, 1.3, 0.07),
    ("BED-006", "Linen Pillowcases (2)", "Bedding", 49.0, 15.0, 1.2, 0.05),
    ("BED-007", "Weighted Throw 15lb", "Bedding", 159.0, 71.0, 0.5, 0.11),
    ("BED-008", "Silk Sleep Mask", "Bedding", 29.0, 9.0, 0.7, 0.04),
    ("BAT-001", "Turkish Cotton Towel Set", "Bath", 89.0, 28.0, 1.6, 0.04),
    ("BAT-002", "Waffle Bath Robe", "Bath", 98.0, 36.0, 0.8, 0.08),
    ("BAT-003", "Teak Bath Mat", "Bath", 69.0, 31.0, 0.5, 0.05),
    ("BAT-004", "Ceramic Soap Dispenser", "Bath", 28.0, 9.5, 0.9, 0.10),
    ("BAT-005", "Hand Towels (2)", "Bath", 32.0, 9.0, 1.3, 0.03),
    ("BAT-006", "Shower Curtain - Linen", "Bath", 59.0, 20.0, 0.6, 0.06),
    ("DEC-001", "Hand-Thrown Vase - Large", "Decor", 85.0, 33.0, 0.6, 0.07),
    ("DEC-002", "Wool Throw Pillow Cover", "Decor", 45.0, 14.0, 1.1, 0.05),
    ("DEC-003", "Soy Candle - Cedar", "Decor", 32.0, 8.0, 2.0, 0.02),
    ("DEC-004", "Soy Candle - Fig", "Decor", 32.0, 8.0, 1.6, 0.02),
    ("DEC-005", "Woven Storage Basket", "Decor", 54.0, 19.0, 0.9, 0.05),
    ("DEC-006", "Framed Botanical Print", "Decor", 76.0, 29.0, 0.4, 0.12),
    ("DEC-007", "Jute Area Rug 5x8", "Decor", 249.0, 118.0, 0.35, 0.14),
    ("DEC-008", "Brass Picture Frame", "Decor", 38.0, 17.0, 0.3, 0.06),
    ("LGT-001", "Ceramic Table Lamp", "Lighting", 139.0, 52.0, 0.6, 0.09),
    ("LGT-002", "Rattan Pendant Shade", "Lighting", 119.0, 47.0, 0.4, 0.10),
    ("LGT-003", "Linen Lamp Shade", "Lighting", 49.0, 17.0, 0.5, 0.06),
    ("LGT-004", "Beeswax Taper Candles (6)", "Lighting", 26.0, 7.5, 1.0, 0.02),
    ("LGT-005", "Glass Hurricane Lantern", "Lighting", 44.0, 18.0, 0.25, 0.15),
    ("LGT-006", "Smart Bulb Starter Kit", "Lighting", 59.0, 41.0, 0.3, 0.08),
]
SKUS = [c[0] for c in CATALOG]
POP = np.array([c[5] for c in CATALOG])
POP = POP / POP.sum()
BY_SKU = {c[0]: c for c in CATALOG}

# --------------------------------------------------------------------------- calendar
days = [START + dt.timedelta(days=i) for i in range((END - START).days + 1)]


def seasonality(d: dt.date) -> float:
    base = {1: 0.80, 2: 0.78, 3: 0.88, 4: 0.92, 5: 1.00, 6: 0.95, 7: 0.90,
            8: 0.95, 9: 1.00, 10: 1.10, 11: 1.55, 12: 1.45}[d.month]
    # BFCM week spike
    if d.month == 11 and 24 <= d.day <= 30:
        base *= 1.9
    weekday = [1.05, 1.0, 1.0, 1.0, 0.95, 0.95, 1.05][d.weekday()]
    return base * weekday


def growth(d: dt.date) -> float:
    # brand grows ~25% over the 18 months
    t = (d - START).days / (END - START).days
    return 1.0 + 0.25 * t


def is_bfcm(d: dt.date) -> bool:
    return d.month == 11 and 24 <= d.day <= 30


# --------------------------------------------------------------------------- ad spend
spend_rows = []
new_customers_plan = []  # (date, channel, n_new)
for d in days:
    t = (d - START).days / (END - START).days
    s = seasonality(d) * growth(d)
    # Meta: spend scaled up aggressively over time while efficiency decays (CAC 38 -> 66)
    meta_spend = 225 * s * (1.0 + 0.55 * t) * rng.lognormal(0, 0.12)
    meta_cac = (46 + 42 * t) * rng.lognormal(0, 0.10) * (0.85 if is_bfcm(d) else 1.0)
    # Google: steadier, mostly branded + shopping (CAC ~44 flat)
    google_spend = 158 * s * (1.0 + 0.10 * t) * rng.lognormal(0, 0.10)
    google_cac = 54 * rng.lognormal(0, 0.10) * (0.9 if is_bfcm(d) else 1.0)

    for channel, total, cac, campaigns in (
        ("meta", meta_spend, meta_cac, [("Meta - Prospecting", 0.72), ("Meta - Retargeting", 0.28)]),
        ("google", google_spend, google_cac, [("Google - Shopping", 0.55), ("Google - Brand Search", 0.25),
                                              ("Google - Non-brand Search", 0.20)]),
    ):
        for camp, share in campaigns:
            amt = round(total * share * rng.lognormal(0, 0.05), 2)
            impressions = int(amt * rng.uniform(60, 110))
            clicks = int(impressions * rng.uniform(0.008, 0.02))
            spend_rows.append([d.isoformat(), channel, camp, f"{amt:.2f}", impressions, clicks])
        new_customers_plan.append((d, channel, rng.poisson(total / cac)))

    # unpaid acquisition
    for channel, lam in (("organic", 2.4), ("direct", 1.5), ("email", 0.45), ("referral", 0.35)):
        new_customers_plan.append((d, channel, rng.poisson(lam * s)))

# --------------------------------------------------------------------------- customers & orders
FIRST = ["alex", "sam", "jordan", "taylor", "casey", "morgan", "riley", "jamie", "quinn", "avery",
         "drew", "rowan", "harper", "logan", "parker", "reese", "skyler", "emerson", "finley", "hayden"]
DOMAINS = ["example.com", "example.org", "mail.example.net", "inbox.example.com"]


def random_time_on(d: dt.date) -> dt.datetime:
    # store-local time, evening-heavy
    hour = int(np.clip(rng.normal(15, 4.5), 0, 23))
    local = dt.datetime(d.year, d.month, d.day, hour, int(rng.integers(0, 60)), int(rng.integers(0, 60)),
                        tzinfo=STORE_TZ)
    return local


orders = []  # dicts at order level
cust_counter = 0
order_counter = 1000
for d, channel, n in new_customers_plan:
    for _ in range(n):
        cust_counter += 1
        email = f"{FIRST[cust_counter % len(FIRST)]}.{cust_counter:05d}@{DOMAINS[cust_counter % 4]}"
        country = "CA" if rng.random() < 0.22 else "US"
        # first-order category affinity drives retention (bedding buyers come back more)
        first_skus = list(rng.choice(SKUS, size=int(rng.choice([1, 1, 2, 2, 3, 3])), replace=False, p=POP))
        bedding_first = any(BY_SKU[s][2] == "Bedding" for s in first_skus)
        repeat_p = 0.36 if bedding_first else 0.22
        if channel in ("email", "referral"):
            repeat_p += 0.08
        # build this customer's order dates
        dates = [d]
        cur = d
        p = repeat_p
        while rng.random() < p:
            gap = int(rng.gamma(2.0, 45)) + 7
            cur = cur + dt.timedelta(days=gap)
            if cur > END:
                break
            dates.append(cur)
            p = min(0.55, p + 0.12)  # loyal customers keep coming back
        for i, od in enumerate(dates):
            order_counter += int(rng.integers(1, 3))
            if i == 0:
                skus, ch = first_skus, channel
            else:
                skus = list(rng.choice(SKUS, size=int(rng.choice([1, 1, 2, 2, 3])), replace=False, p=POP))
                ch = rng.choice(["email", "direct", "organic", "meta", "google"], p=[0.38, 0.25, 0.17, 0.10, 0.10])
            orders.append(dict(order_id=order_counter, email=email, country=country, date=od,
                               ts=random_time_on(od), channel=str(ch), skus=skus, new=(i == 0)))

orders.sort(key=lambda o: o["ts"])

# FX
fx = {}
m = dt.date(START.year, START.month, 1)
rate = 0.725
while m <= END:
    rate = rate + rng.normal(0, 0.006)
    fx[m.strftime("%Y-%m")] = round(rate, 4)
    m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)

# --------------------------------------------------------------------------- line items
DISCOUNT_CODES = {"WELCOME10": 0.10, "BFCM25": 0.25, "VIP15": 0.15, "SPRING20": 0.20}
header = ["order_id", "line_number", "created_at", "customer_email", "billing_country", "currency",
          "sku", "product_title", "quantity", "unit_price", "line_discount", "discount_code",
          "refunded_amount", "financial_status", "source_channel"]
lines = []
for o in orders:
    cad = o["country"] == "CA"
    code = ""
    if is_bfcm(o["date"]):
        code = "BFCM25" if rng.random() < 0.8 else ""
    elif o["new"] and rng.random() < 0.30:
        code = "WELCOME10"
    elif not o["new"] and rng.random() < 0.20:
        code = "VIP15"
    elif o["date"].month in (3, 4) and rng.random() < 0.15:
        code = "SPRING20"
    pct = DISCOUNT_CODES.get(code, 0.0)

    status = "paid"
    r = rng.random()
    if r < 0.006:
        status = "voided"
    order_refund_roll = rng.random()

    # timestamp format mess: most UTC 'Z', some local offset, a few naive (legacy export, store-local)
    f = rng.random()
    if f < 0.70:
        ts = o["ts"].astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    elif f < 0.95:
        ts = o["ts"].isoformat()
    else:
        ts = o["ts"].strftime("%Y-%m-%d %H:%M:%S")

    for ln, sku in enumerate(o["skus"], start=1):
        _, title, cat, price_usd, _, _, refund_p = BY_SKU[sku]
        qty = int(rng.choice([1, 1, 1, 1, 1, 2]))
        unit = round(price_usd / fx[o["date"].strftime("%Y-%m")] * 1.0, 2) if cad else price_usd
        unit = round(math.ceil(unit) - 0.01, 2) if cad else unit  # CAD price points end in .99
        disc = round(unit * qty * pct, 2)
        refund = 0.0
        if status == "paid" and (rng.random() < refund_p or order_refund_roll < 0.01):
            refund = round(unit * qty - disc, 2) if rng.random() < 0.8 else round((unit * qty - disc) * 0.5, 2)
        fin = status
        if status == "paid" and refund > 0:
            fin = "partially_refunded"
        lines.append([o["order_id"], ln, ts, o["email"], o["country"], "CAD" if cad else "USD", sku, title,
                      qty, f"{unit:.2f}", f"{disc:.2f}", code, f"{refund:.2f}", fin, o["channel"]])

# --------------------------------------------------------------------------- inject dirt
n = len(lines)
# 1) null emails (guest checkouts / export glitch)
for i in rng.choice(n, size=142, replace=False):
    lines[i][3] = ""
# 2) negative quantities (manual adjustments keyed into the wrong field)
neg_idx = rng.choice(n, size=37, replace=False)
for i in neg_idx:
    lines[i][8] = -abs(int(lines[i][8]))
# 3) test orders from staff
for k in range(12):
    d = START + dt.timedelta(days=int(rng.integers(0, 540)))
    lines.append([900000 + k, 1, f"{d.isoformat()}T12:00:00Z", f"qa+{k}@northwind-test.example", "US", "USD",
                  "DEC-003", "Soy Candle - Cedar", 1, "32.00", "0.00", "", "0.00", "paid", "direct"])
# 4) duplicated order rows (export pagination overlap): whole orders re-exported
order_ids = sorted({l[0] for l in lines if l[0] < 900000})
dupe_orders = set(int(x) for x in rng.choice(order_ids, size=85, replace=False))
dupes = [list(l) for l in lines if l[0] in dupe_orders]
lines.extend(dupes)
# 5) whitespace / casing noise on currency + email
for i in rng.choice(len(lines), size=60, replace=False):
    lines[i][5] = lines[i][5].lower() + " "
for i in rng.choice(len(lines), size=40, replace=False):
    if lines[i][3]:
        lines[i][3] = "  " + lines[i][3].upper()

# shuffle export order a bit (exports are not guaranteed sorted)
perm = rng.permutation(len(lines))
lines = [lines[i] for i in perm]

# 6) ad-spend export overlap: Google re-exported 3 days (duplicate rows)
g_dupe_days = {"2025-08-14", "2026-01-02", "2026-05-20"}
spend_rows.extend([list(r) for r in spend_rows if r[1] == "google" and r[0] in g_dupe_days])

# --------------------------------------------------------------------------- write
RAW.mkdir(parents=True, exist_ok=True)
with open(RAW / "shopify_orders_export.csv", "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(header)
    w.writerows(lines)
with open(RAW / "ad_spend_daily.csv", "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["date", "channel", "campaign", "spend_usd", "impressions", "clicks"])
    w.writerows(spend_rows)
with open(RAW / "products.csv", "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["sku", "product_title", "category", "list_price_usd", "unit_cost_usd"])
    for c in CATALOG:
        w.writerow([c[0], c[1], c[2], f"{c[3]:.2f}", f"{c[4]:.2f}"])
with open(RAW / "fx_rates_monthly.csv", "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["month", "cad_to_usd"])
    for k, v in fx.items():
        w.writerow([k, f"{v:.4f}"])

print(f"orders export rows: {len(lines):,}  (unique order ids: {len({l[0] for l in lines}):,})")
print(f"ad spend rows: {len(spend_rows):,}")
print(f"customers simulated: {cust_counter:,}")
