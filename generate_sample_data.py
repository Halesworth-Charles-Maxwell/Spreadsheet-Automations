#!/usr/bin/env python3
"""
generate_sample_data.py

Creates four deliberately messy exports for a FICTIONAL store, "Northwind Candle Co.":

    sample_data/shopify_orders.csv   - Shopify-style order export
    sample_data/amazon_report.csv    - Amazon-style transaction report
    sample_data/ad_spend.csv         - Daily ad spend (Meta, Google, Amazon Ads)
    sample_data/product_costs.xlsx   - Unit costs per SKU

All data is randomly generated (seeded, so every run produces the same files).
Nothing here comes from a real business.

Usage:
    python generate_sample_data.py
"""
import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

SEED = 42
START, END = date(2026, 4, 1), date(2026, 9, 30)
OUT_DIR = Path(__file__).parent / "sample_data"

# sku, name, price, unit cost, Shopify popularity, Amazon popularity
PRODUCTS = [
    ("CND-01", "Lavender Fields 8oz Candle", 24.00, 7.50, 14, 12),
    ("CND-02", "Vanilla Bean 8oz Candle", 24.00, 7.20, 16, 14),
    ("CND-03", "Cedar & Sage 12oz Candle", 32.00, 10.50, 12, 8),
    ("CND-04", "Ocean Breeze 8oz Candle", 22.00, 8.00, 8, 10),
    ("CND-05", "Pumpkin Spice 8oz Candle", 26.00, 9.00, 10, 8),
    ("CND-06", "Sandalwood 12oz Candle", 34.00, 12.00, 9, 6),
    ("CND-07", "Wax Melts 6-Pack", 14.00, 8.20, 7, 16),
    ("CND-08", "Candle Gift Set (3 x 4oz)", 58.00, 27.00, 11, 4),
    ("CND-09", "Travel Tin 4oz Candle", 12.00, 5.20, 6, 9),
    ("CND-10", "Reed Diffuser 100ml", 29.00, 11.00, 7, 6),
]
PRICE = {p[0]: p[2] for p in PRODUCTS}
NAME = {p[0]: p[1] for p in PRODUCTS}

DAYS = (END - START).days + 1


# ---------- helpers that create realistic mess ----------
def seasonal(sku, d):
    """Pumpkin Spice sells in autumn; gift sets spike in September."""
    if sku == "CND-05":
        return {4: 0.2, 5: 0.2, 6: 0.2, 7: 0.3, 8: 1.0, 9: 3.0}.get(d.month, 1.0)
    if sku == "CND-08":
        return 1.6 if d.month == 9 else 1.0
    return 1.0


def pick_skus(d, channel, k):
    idx = 4 if channel == "shopify" else 5
    pool = [(p[0], p[idx] * seasonal(p[0], d)) for p in PRODUCTS]
    chosen = []
    for _ in range(k):
        sku = random.choices([x[0] for x in pool], weights=[x[1] for x in pool])[0]
        chosen.append(sku)
        pool = [x for x in pool if x[0] != sku]
    return chosen


def messy_sku(sku):
    r = random.random()
    if r < 0.05:
        return sku.lower()
    if r < 0.09:
        return sku + " "
    if r < 0.12:
        return sku.lower().replace("-", "_")
    if r < 0.14:
        return " " + sku
    return sku


def money_text(x):
    """Mostly plain numbers, sometimes with a $ symbol (stored as text)."""
    return f"${x:,.2f}" if random.random() < 0.06 else f"{x:.2f}"


def growth(d):
    return 1.0 + 0.25 * ((d - START).days / DAYS)


def daterange():
    for i in range(DAYS):
        yield START + timedelta(days=i)


# ---------- Shopify ----------
def build_shopify():
    header = ["Name", "Created at", "Lineitem name", "Lineitem sku",
              "Lineitem quantity", "Lineitem price", "Discount Amount",
              "Refunded Amount"]
    rows, order_no = [], 1000
    for d in daterange():
        n_orders = max(1, round(random.uniform(3.5, 7.5) * growth(d) * (1.15 if d.weekday() >= 5 else 1)))
        for _ in range(n_orders):
            order_no += 1
            ts = datetime(d.year, d.month, d.day, random.randint(7, 23),
                          random.randint(0, 59), random.randint(0, 59))
            created = ts.strftime("%Y-%m-%d %H:%M:%S") + " -0400"
            k = random.choices([1, 2, 3], weights=[70, 25, 5])[0]
            has_code = random.random() < 0.15
            for sku in pick_skus(d, "shopify", k):
                qty = random.choice([1, 1, 1, 2, 2, 3])
                line_total = qty * PRICE[sku]
                discount = round(0.10 * line_total, 2) if has_code else 0.0
                refund = round(line_total - discount, 2) if random.random() < 0.04 else 0.0
                rows.append([f"#{order_no}", created, NAME[sku], messy_sku(sku), str(qty),
                             money_text(PRICE[sku]), f"{discount:.2f}", f"{refund:.2f}"])

    # Mess 1: three rows lose their SKU
    for i in random.sample(range(len(rows)), 3):
        rows[i][3] = ""
    missing = {i for i, r in enumerate(rows) if r[3] == ""}

    # Mess 2: exact duplicate rows
    dup_idx = random.sample([i for i in range(len(rows)) if i not in missing], 14)
    for i in sorted(dup_idx, reverse=True):
        rows.insert(i + 1, list(rows[i]))

    # Mess 3: blank rows (all commas)
    for pos in sorted(random.sample(range(len(rows)), 5), reverse=True):
        rows.insert(pos, [""] * len(header))
    return header, rows


# ---------- Amazon ----------
def build_amazon():
    header = ["amazon-order-id", "purchase-date", "transaction-type", "sku",
              "quantity-purchased", "item-price", "promotion-discount", "amazon-fees"]
    entries, used_ids = [], set()
    for d in daterange():
        n_orders = max(1, round(random.uniform(2, 5) * growth(d)))
        for _ in range(n_orders):
            while True:
                oid = f"113-{random.randint(1000000, 9999999)}-{random.randint(1000000, 9999999)}"
                if oid not in used_ids:
                    used_ids.add(oid)
                    break
            ts = datetime(d.year, d.month, d.day, random.randint(0, 23), random.randint(0, 59))
            sku = pick_skus(d, "amazon", 1)[0]
            qty = random.choice([1, 1, 1, 2])
            total = qty * PRICE[sku]  # NOTE: Amazon item-price is the LINE total
            promo = 2.00 if random.random() < 0.10 else 0.0
            fees = round(0.15 * total + 3.20 * qty, 2)
            entries.append((ts, [oid, d.strftime("%m/%d/%Y"), "Order", messy_sku(sku), str(qty),
                                 money_text(total), f"{promo:.2f}", money_text(fees)]))
            if random.random() < 0.05:  # refunds arrive later as NEGATIVE rows
                rd = d + timedelta(days=random.randint(3, 20))
                refund_amt = total - promo
                neg = f"-${refund_amt:,.2f}" if random.random() < 0.5 else f"-{refund_amt:.2f}"
                entries.append((datetime(rd.year, rd.month, rd.day, 12, 0),
                                [oid, rd.strftime("%m/%d/%Y"), "Refund", messy_sku(sku), str(qty),
                                 neg, "0.00", "0.00"]))
    entries.sort(key=lambda e: e[0])
    rows = [e[1] for e in entries]

    for i in sorted(random.sample(range(len(rows)), 8), reverse=True):
        rows.insert(i + 1, list(rows[i]))
    return header, rows


# ---------- Ad spend ----------
def build_ads():
    header = ["Date", "Platform", "Campaign", "Spend"]
    platforms = [
        ("Meta Ads", ["Retargeting", "Prospecting - Lookalike"], 45),
        ("Google Ads", ["Brand Search", "Shopping"], 25),
        ("Amazon Ads", ["Sponsored Products", "Sponsored Brands"], 28),
    ]
    rows = []
    for d in daterange():
        for name, campaigns, base in platforms:
            for camp in campaigns:
                spend = base / 2 * random.uniform(0.6, 1.4) * growth(d)
                r = random.random()
                label = name.upper() if r < 0.04 else name.lower() if r < 0.08 else name
                rows.append([d.isoformat(), label, camp, money_text(spend)])
    return header, rows


# ---------- Product costs ----------
def build_costs():
    suppliers = ["Harbor Wax Supply", "Bluebird Packaging", "Aroma Source Ltd."]
    records = []
    for sku, name, _price, cost, _s, _a in PRODUCTS:
        if sku == "CND-10":
            continue  # MESS: this SKU sells, but is missing from the cost file
        sku_out = {"CND-04": "cnd-04", "CND-02": "CND-02 "}.get(sku, sku)
        cost_out = f"${cost:.2f}" if sku == "CND-06" else cost  # one cost stored as text
        records.append([sku_out, name, cost_out, random.choice(suppliers)])
    records.append(["CND-11", "Holiday Wreath Candle (discontinued)", 15.00, suppliers[0]])
    return pd.DataFrame(records, columns=["SKU", "Product Name", "Unit Cost", "Supplier"])


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    random.seed(SEED)
    OUT_DIR.mkdir(exist_ok=True)

    for filename, builder in [("shopify_orders.csv", build_shopify),
                              ("amazon_report.csv", build_amazon),
                              ("ad_spend.csv", build_ads)]:
        header, rows = builder()
        write_csv(OUT_DIR / filename, header, rows)
        print(f"Wrote {filename:<20} {len(rows):>5} rows")

    build_costs().to_excel(OUT_DIR / "product_costs.xlsx", index=False)
    print(f"Wrote {'product_costs.xlsx':<20} {len(build_costs()):>5} rows")
    print(f"\nSample data saved to: {OUT_DIR}")


if __name__ == "__main__":
    main()
