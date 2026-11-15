#!/usr/bin/env python3
"""
build_report.py

Cleans and merges four messy exports (Shopify orders, Amazon transactions, ad spend,
product costs) into one formatted Excel profit report.

Usage:
    python build_report.py
    python build_report.py --input-dir sample_data --output northwind_profit_report.xlsx

Output workbook:
    Summary       - key figures, channel breakdown, best/worst products, charts
    By Product    - profit and margin per SKU (before ad spend)
    By Month      - revenue, costs, ad spend and profit per month
    Data Quality  - everything the script found and fixed, plus rows needing review
    Orders        - the clean, merged order data (the report's formulas read from here)
    Ad Spend      - the clean ad data
    Notes         - definitions and assumptions

Every number on the report sheets is an Excel formula reading from Orders / Ad Spend,
so the workbook stays live if someone edits the clean data.
"""
from __future__ import annotations

import argparse
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill

# =====================================================================================
# SETTINGS - change these for a real client
# =====================================================================================
STORE_NAME = "Northwind Candle Co."
SHOPIFY_FEE_PCT = 0.029      # payment processing, % of line revenue
SHOPIFY_FEE_FIXED = 0.30     # payment processing, $ per order
LOW_MARGIN_THRESHOLD = 0.20  # margins below this are highlighted on "By Product"
AD_PLATFORM_TO_CHANNEL = {   # which sales channel each ad platform is charged to
    "META ADS": "Shopify",
    "GOOGLE ADS": "Shopify",
    "AMAZON ADS": "Amazon",
}
PLATFORM_DISPLAY = {"META ADS": "Meta Ads", "GOOGLE ADS": "Google Ads", "AMAZON ADS": "Amazon Ads"}
CHANNELS = ["Shopify", "Amazon"]

FONT = "Arial"
NAVY, LIGHT, YELLOW = "1F3864", "D9E1F2", "FFFF00"
CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00'
PCT = '0.0%'


# =====================================================================================
# Data-quality log
# =====================================================================================
@dataclass
class QualityLog:
    counts: list = field(default_factory=list)     # (file, rows read, rows used)
    fixes: list = field(default_factory=list)      # (source, issue, rows, action)
    attention: list = field(default_factory=list)  # (source, sku, reason, reference)

    def fix(self, source, issue, rows, action):
        if rows:
            self.fixes.append((source, issue, int(rows), action))

    def flag(self, source, sku, reason, reference=""):
        self.attention.append((source, sku, reason, reference))


def _txt(v) -> str:
    return "" if pd.isna(v) else str(v)


# =====================================================================================
# Cleaning helpers
# =====================================================================================
def clean_sku(s: pd.Series) -> pd.Series:
    """' cnd_03 ' -> 'CND-03'"""
    return s.str.strip().str.upper().str.replace("_", "-", regex=False)


def to_number(s: pd.Series) -> pd.Series:
    """'$1,234.50' -> 1234.5, '-$24.00' -> -24.0, junk -> NaN"""
    return pd.to_numeric(s.str.replace(r"[$,\s]", "", regex=True), errors="coerce")


def count_symbols(s: pd.Series) -> int:
    return int(s.str.contains(r"[$,]", na=False).sum())


def drop_blank_and_duplicate_rows(df, log, source):
    blank = df.isna().all(axis=1)
    log.fix(source, "Blank rows", blank.sum(), "Removed")
    df = df[~blank]
    dup = df.duplicated()
    log.fix(source, "Exact duplicate rows", dup.sum(), "Removed (kept first copy)")
    return df[~dup].copy()


def drop_and_flag(df, mask, log, source, reason, ref_col, sku_col="SKU"):
    """Drop rows matching mask and list them on the Data Quality sheet."""
    for _, r in df[mask].iterrows():
        sku = _txt(r[sku_col]) if sku_col in df.columns else ""
        log.flag(source, sku, reason, _txt(r[ref_col]))
    log.fix(source, reason, mask.sum(), "Dropped and listed under 'Rows needing your attention'")
    return df[~mask].copy()


# =====================================================================================
# Loaders - one per source file. Each returns a standard table.
# =====================================================================================
ORDER_COLS = ["Order ID", "Date", "Channel", "SKU", "Qty", "Revenue", "Refund", "Fees", "Product"]


def load_shopify(path: Path, log: QualityLog) -> pd.DataFrame:
    src = path.name
    raw = pd.read_csv(path, dtype=str)
    df = drop_blank_and_duplicate_rows(raw, log, src)

    raw_sku = df["Lineitem sku"]
    df["SKU"] = clean_sku(raw_sku)
    df = drop_and_flag(df, df["SKU"].isna() | (df["SKU"] == ""), log, src,
                       "Rows with missing SKU", "Name")
    log.fix(src, "SKUs with stray spaces / lowercase / underscores",
            (raw_sku.loc[df.index] != df["SKU"]).sum(), "Standardized (e.g. ' cnd_03 ' -> 'CND-03')")

    # Dates look like '2026-04-01 07:47:17 -0400'. Use the date as shown (store time).
    df["Date"] = pd.to_datetime(df["Created at"].str[:10], format="%Y-%m-%d", errors="coerce")
    df = drop_and_flag(df, df["Date"].isna(), log, src, "Rows with unreadable date", "Name")

    log.fix(src, "Prices stored as text with $ or commas", count_symbols(df["Lineitem price"]),
            "Symbols stripped, converted to numbers")
    qty = to_number(df["Lineitem quantity"])
    price = to_number(df["Lineitem price"])
    df = drop_and_flag(df, qty.isna() | price.isna(), log, src, "Rows with unreadable quantity/price", "Name")
    qty, price = qty.loc[df.index], price.loc[df.index]

    discount = to_number(df["Discount Amount"]).fillna(0)
    refund = to_number(df["Refunded Amount"]).fillna(0)

    out = pd.DataFrame({
        "Order ID": df["Name"],
        "Date": df["Date"],
        "Channel": "Shopify",
        "SKU": df["SKU"],
        "Qty": qty.astype(int),
        "Revenue": (qty * price - discount).round(2),
        "Refund": refund.abs().round(2),
        "Product": df["Lineitem name"],
    })
    first_line_of_order = ~out["Order ID"].duplicated()  # fixed fee charged once per order
    out["Fees"] = (SHOPIFY_FEE_PCT * out["Revenue"] + SHOPIFY_FEE_FIXED * first_line_of_order).round(2)

    log.counts.append((src, len(raw), len(out)))
    return out[ORDER_COLS]


def load_amazon(path: Path, log: QualityLog) -> pd.DataFrame:
    src = path.name
    raw = pd.read_csv(path, dtype=str)
    df = drop_blank_and_duplicate_rows(raw, log, src)

    raw_sku = df["sku"]
    df["SKU"] = clean_sku(raw_sku)
    df = drop_and_flag(df, df["SKU"].isna() | (df["SKU"] == ""), log, src,
                       "Rows with missing SKU", "amazon-order-id")
    log.fix(src, "SKUs with stray spaces / lowercase / underscores",
            (raw_sku.loc[df.index] != df["SKU"]).sum(), "Standardized (e.g. ' cnd_03 ' -> 'CND-03')")

    is_refund = df["transaction-type"].str.strip().str.lower() == "refund"
    orders, refunds = df[~is_refund].copy(), df[is_refund].copy()

    # US-style dates: 04/01/2026 = April 1 (parsed explicitly so day/month never get swapped)
    orders["Date"] = pd.to_datetime(orders["purchase-date"], format="%m/%d/%Y", errors="coerce")
    orders = drop_and_flag(orders, orders["Date"].isna(), log, src, "Rows with unreadable date",
                           "amazon-order-id")

    log.fix(src, "Amounts stored as text with $ or commas",
            sum(count_symbols(orders[c]) for c in ["item-price", "amazon-fees"]),
            "Symbols stripped, converted to numbers")
    qty = to_number(orders["quantity-purchased"])
    line_total = to_number(orders["item-price"])  # Amazon's item-price is the LINE total, not per unit
    orders = drop_and_flag(orders, qty.isna() | line_total.isna(), log, src,
                           "Rows with unreadable quantity/price", "amazon-order-id")
    qty, line_total = qty.loc[orders.index], line_total.loc[orders.index]
    promo = to_number(orders["promotion-discount"]).fillna(0)
    fees = to_number(orders["amazon-fees"]).abs().fillna(0)

    out = pd.DataFrame({
        "Order ID": orders["amazon-order-id"],
        "Date": orders["Date"],
        "Channel": "Amazon",
        "SKU": orders["SKU"],
        "Qty": qty.astype(int),
        "Revenue": (line_total - promo).round(2),
        "Fees": fees.round(2),
        "Product": pd.NA,
    })

    # Refunds arrive as separate NEGATIVE rows: attach them to the original order
    log.fix(src, "Refunds recorded as separate negative rows", len(refunds),
            "Converted into a Refund column on the original order")
    refunds["Refund"] = to_number(refunds["item-price"]).abs()
    refund_by_order = refunds.groupby(["amazon-order-id", "SKU"])["Refund"].sum()
    keys = list(zip(out["Order ID"], out["SKU"]))
    out["Refund"] = [round(float(refund_by_order.get(k, 0.0)), 2) for k in keys]
    for (oid, sku), amt in refund_by_order.items():
        if (oid, sku) not in set(keys):
            log.flag(src, sku, f"Refund of ${amt:,.2f} has no matching order - not counted", oid)

    log.counts.append((src, len(raw), len(out)))
    return out[ORDER_COLS]


def load_ads(path: Path, log: QualityLog) -> pd.DataFrame:
    src = path.name
    raw = pd.read_csv(path, dtype=str)
    df = drop_blank_and_duplicate_rows(raw, log, src)

    key = df["Platform"].str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    df["PlatformKey"] = key
    log.fix(src, "Platform names with inconsistent capitalization",
            (df["Platform"] != key.map(PLATFORM_DISPLAY)).sum() - key.map(PLATFORM_DISPLAY).isna().sum(),
            "Standardized (e.g. 'META ADS' -> 'Meta Ads')")
    df = drop_and_flag(df, ~key.isin(AD_PLATFORM_TO_CHANNEL), log, src,
                       "Rows with unrecognised ad platform", "Campaign", sku_col="")

    df["Date"] = pd.to_datetime(df["Date"], format="%Y-%m-%d", errors="coerce")
    log.fix(src, "Spend stored as text with $ or commas", count_symbols(df["Spend"]),
            "Symbols stripped, converted to numbers")
    spend = to_number(df["Spend"])
    df = drop_and_flag(df, df["Date"].isna() | spend.isna(), log, src,
                       "Rows with unreadable date/spend", "Campaign", sku_col="")
    spend = spend.loc[df.index]

    out = pd.DataFrame({
        "Date": df["Date"],
        "Platform": df["PlatformKey"].map(PLATFORM_DISPLAY),
        "Campaign": df["Campaign"],
        "Channel": df["PlatformKey"].map(AD_PLATFORM_TO_CHANNEL),
        "Spend": spend.round(2),
    })
    log.counts.append((src, len(raw), len(out)))
    return out.sort_values(["Date", "Platform", "Campaign"]).reset_index(drop=True)


def load_costs(path: Path, log: QualityLog) -> pd.DataFrame:
    src = path.name
    raw = pd.read_excel(path, dtype=str)
    df = drop_blank_and_duplicate_rows(raw, log, src)
    raw_sku = df["SKU"]
    df["SKU"] = clean_sku(raw_sku)
    log.fix(src, "SKUs with stray spaces / lowercase / underscores", (raw_sku != df["SKU"]).sum(),
            "Standardized so they match the sales data")
    log.fix(src, "Unit costs stored as text with $", count_symbols(df["Unit Cost"]),
            "Symbols stripped, converted to numbers")
    df["Unit Cost"] = to_number(df["Unit Cost"])
    dup_sku = df["SKU"].duplicated()
    log.fix(src, "SKUs listed more than once", dup_sku.sum(), "Kept the first entry")
    df = df[~dup_sku]
    log.counts.append((src, len(raw), len(df)))
    return df[["SKU", "Product Name", "Unit Cost"]]


# =====================================================================================
# Combine
# =====================================================================================
def combine(shopify, amazon, costs, log: QualityLog) -> pd.DataFrame:
    orders = pd.concat([shopify, amazon], ignore_index=True)
    orders = orders.sort_values(["Date", "Channel", "Order ID"]).reset_index(drop=True)
    orders["Month"] = orders["Date"].dt.strftime("%Y-%m")

    # Product names: take the most common name seen in the sales data, fall back to the cost file
    names = orders.dropna(subset=["Product"]).groupby("SKU")["Product"].agg(lambda s: s.mode().iat[0])
    names = names.combine_first(costs.set_index("SKU")["Product Name"])
    orders["Product"] = orders["SKU"].map(names).fillna("(name unavailable)")

    # COGS = units x unit cost. SKUs without a cost get $0 and are flagged, never silently dropped.
    unit_cost = orders["SKU"].map(costs.set_index("SKU")["Unit Cost"])
    for sku in sorted(orders.loc[unit_cost.isna(), "SKU"].unique()):
        n = int((orders["SKU"] == sku).sum())
        log.fix("product_costs.xlsx", f"SKU {sku} sells but has no unit cost", n,
                "COGS set to $0 (profit overstated). Add the cost and re-run.")
        log.flag("product_costs.xlsx", sku, f"No unit cost found ({n} order lines affected)", "")
    orders["COGS"] = (orders["Qty"] * unit_cost.fillna(0)).round(2)

    return orders[["Order ID", "Date", "Month", "Channel", "SKU", "Qty", "Revenue", "Refund",
                   "Fees", "COGS", "Product"]]


# =====================================================================================
# Excel writing helpers
# =====================================================================================
def cell(ws, row, col, value=None, *, bold=False, color="000000", fmt=None, fill=None,
         align=None, size=10, italic=False, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="top", wrap_text=wrap)
    return c


def header(ws, row, col, text):
    return cell(ws, row, col, text, bold=True, color="FFFFFF", fill=NAVY, align="center", wrap=True)


def section(ws, row, text):
    cell(ws, row, 1, text, bold=True, size=12, color=NAVY)


def title(ws, text, subtitle):
    cell(ws, 1, 1, text, bold=True, size=16, color=NAVY)
    cell(ws, 2, 1, subtitle, italic=True, color="595959")


def widths(ws, spec: dict):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def style_chart(chart, title_text, x_title=None, y_title=None, legend=True):
    chart.title = title_text
    chart.title.overlay = False
    chart.roundedCorners = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if x_title:
        chart.x_axis.title = x_title
        chart.x_axis.title.overlay = False
    if y_title:
        chart.y_axis.title = y_title
        chart.y_axis.title.overlay = False
    if legend:
        chart.legend.position = "b"
        chart.legend.overlay = False
    else:
        chart.legend = None


# =====================================================================================
# Workbook
# =====================================================================================
def write_workbook(orders: pd.DataFrame, ads: pd.DataFrame, log: QualityLog, out_path: Path):
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_prod = wb.create_sheet("By Product")
    ws_mon = wb.create_sheet("By Month")
    ws_dq = wb.create_sheet("Data Quality")
    ws_ord = wb.create_sheet("Orders")
    ws_ads = wb.create_sheet("Ad Spend")
    ws_notes = wb.create_sheet("Notes")

    period = f"{orders['Date'].min():%b %d, %Y} to {orders['Date'].max():%b %d, %Y}"
    subtitle = f"{STORE_NAME}  |  {period}  |  Generated {datetime.now():%Y-%m-%d %H:%M}"

    # ---------------- Orders (clean data) ----------------
    heads = ["Order ID", "Date", "Month", "Channel", "SKU", "Qty", "Revenue", "Refunds", "Fees",
             "COGS", "Profit"]
    for j, h in enumerate(heads, 1):
        header(ws_ord, 1, j, h)
    for i, r in enumerate(orders.itertuples(index=False), start=2):
        cell(ws_ord, i, 1, r[0])
        cell(ws_ord, i, 2, r[1].date(), fmt="yyyy-mm-dd")
        cell(ws_ord, i, 3, r[2])
        cell(ws_ord, i, 4, r[3])
        cell(ws_ord, i, 5, r[4])
        cell(ws_ord, i, 6, int(r[5]), fmt="0")
        cell(ws_ord, i, 7, float(r[6]), fmt=CUR2)
        cell(ws_ord, i, 8, float(r[7]), fmt=CUR2)
        cell(ws_ord, i, 9, float(r[8]), fmt=CUR2)
        cell(ws_ord, i, 10, float(r[9]), fmt=CUR2)
        cell(ws_ord, i, 11, f"=G{i}-H{i}-I{i}-J{i}", fmt=CUR2)
    ord_last = len(orders) + 1
    ws_ord.freeze_panes = "A2"
    ws_ord.auto_filter.ref = f"A1:K{ord_last}"
    widths(ws_ord, {"A": 22, "B": 12, "C": 10, "D": 10, "E": 10, "F": 7, "G": 12, "H": 12,
                    "I": 12, "J": 12, "K": 12})

    def ORD(col):
        return f"Orders!${col}$2:${col}${ord_last}"

    # ---------------- Ad Spend (clean data) ----------------
    for j, h in enumerate(["Date", "Month", "Platform", "Campaign", "Channel", "Spend"], 1):
        header(ws_ads, 1, j, h)
    for i, r in enumerate(ads.itertuples(index=False), start=2):
        cell(ws_ads, i, 1, r[0].date(), fmt="yyyy-mm-dd")
        cell(ws_ads, i, 2, r[0].strftime("%Y-%m"))
        cell(ws_ads, i, 3, r[1])
        cell(ws_ads, i, 4, r[2])
        cell(ws_ads, i, 5, r[3])
        cell(ws_ads, i, 6, float(r[4]), fmt=CUR2)
    ads_last = len(ads) + 1
    ws_ads.freeze_panes = "A2"
    ws_ads.auto_filter.ref = f"A1:F{ads_last}"
    widths(ws_ads, {"A": 12, "B": 10, "C": 14, "D": 26, "E": 10, "F": 12})

    def ADS(col):
        return f"'Ad Spend'!${col}$2:${col}${ads_last}"

    # ---------------- By Product ----------------
    prod = orders.groupby(["SKU", "Product"], as_index=False).agg(
        rev=("Revenue", "sum"), ref=("Refund", "sum"), fee=("Fees", "sum"), cogs=("COGS", "sum"))
    prod["margin"] = (prod.rev - prod.ref - prod.fee - prod.cogs) / (prod.rev - prod.ref)
    missing_cost = set(s[1] for s in log.attention if s[0] == "product_costs.xlsx")
    # Best margin first; SKUs with missing costs go last because their margin is not real
    prod["no_cost"] = prod["SKU"].isin(missing_cost)
    prod = prod.sort_values(["no_cost", "margin"], ascending=[True, False]).reset_index(drop=True)

    title(ws_prod, "Profit by Product (before ad spend)",
          "Ad spend is tracked by channel and month, not by product, so product profit excludes it.")
    cell(ws_prod, 2, 8, "Flag margins below:", bold=True, align="right")
    cell(ws_prod, 2, 9, LOW_MARGIN_THRESHOLD, fmt="0%", color="0000FF", fill=YELLOW, bold=True)

    p_heads = ["SKU", "Product", "Units", "Revenue", "Refunds", "Fees", "COGS",
               "Profit (before ads)", "Margin (before ads)", "Note", "Profit (ranked)"]
    for j, h in enumerate(p_heads, 1):
        header(ws_prod, 4, j, h)
    p_first, p_last = 5, 4 + len(prod)
    for i, r in enumerate(prod.itertuples(index=False), start=p_first):
        cell(ws_prod, i, 1, r.SKU)
        cell(ws_prod, i, 2, r.Product)
        cell(ws_prod, i, 3, f"=SUMIFS({ORD('F')},{ORD('E')},$A{i})", fmt="#,##0")
        for col, src_col in zip(range(4, 8), ["G", "H", "I", "J"]):
            cell(ws_prod, i, col, f"=SUMIFS({ORD(src_col)},{ORD('E')},$A{i})", fmt=CUR)
        cell(ws_prod, i, 8, f"=D{i}-E{i}-F{i}-G{i}", fmt=CUR, bold=True)
        cell(ws_prod, i, 9, f"=IFERROR(H{i}/(D{i}-E{i}),0)", fmt=PCT)
        cell(ws_prod, i, 10, "No unit cost - profit overstated" if r.SKU in missing_cost else None,
             color="9C5700", italic=True)
        cell(ws_prod, i, 11, f'=IF(J{i}="",H{i},"")', fmt=CUR, color="808080")
    t = p_last + 1
    cell(ws_prod, t, 1, "Total", bold=True, fill=LIGHT)
    cell(ws_prod, t, 2, None, fill=LIGHT)
    for col in range(3, 9):
        L = chr(64 + col)
        cell(ws_prod, t, col, f"=SUM({L}{p_first}:{L}{p_last})", bold=True, fill=LIGHT,
             fmt="#,##0" if col == 3 else CUR)
    cell(ws_prod, t, 9, f"=IFERROR(H{t}/(D{t}-E{t}),0)", bold=True, fill=LIGHT, fmt=PCT)
    cell(ws_prod, t, 10, None, fill=LIGHT)
    cell(ws_prod, t, 11, None, fill=LIGHT)

    red = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")
    ws_prod.conditional_formatting.add(
        f"I{p_first}:I{p_last}", CellIsRule(operator="lessThan", formula=["$I$2"], fill=red,
                                            font=Font(name=FONT, color="9C0006")))
    ws_prod.conditional_formatting.add(
        f"H{p_first}:H{p_last}", CellIsRule(operator="lessThan", formula=["0"], fill=red,
                                            font=Font(name=FONT, color="9C0006")))
    ws_prod.conditional_formatting.add(
        f"A{p_first}:J{p_last}", FormulaRule(formula=[f'$J{p_first}<>""'],
                                             fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
    ws_prod.freeze_panes = "C5"
    widths(ws_prod, {"A": 10, "B": 30, "C": 9, "D": 13, "E": 12, "F": 12, "G": 13, "H": 18,
                     "I": 18, "J": 32, "K": 15})
    ws_prod.row_dimensions[4].height = 30

    def PROD(col):
        return f"'By Product'!${col}${p_first}:${col}${p_last}"

    # ---------------- By Month ----------------
    months = sorted(set(orders["Month"]) | set(ads["Date"].dt.strftime("%Y-%m")))
    title(ws_mon, "Monthly Performance", "Refunds are counted in the month of the original order.")
    m_heads = ["Month", "Revenue", "Refunds", "Fees", "COGS", "Profit before ads", "Ad spend",
               "Profit after ads", "Margin after ads"]
    for j, h in enumerate(m_heads, 1):
        header(ws_mon, 4, j, h)
    m_first, m_last = 5, 4 + len(months)
    for i, m in enumerate(months, start=m_first):
        cell(ws_mon, i, 1, m)
        for col, src_col in zip(range(2, 6), ["G", "H", "I", "J"]):
            cell(ws_mon, i, col, f"=SUMIFS({ORD(src_col)},{ORD('C')},$A{i})", fmt=CUR)
        cell(ws_mon, i, 6, f"=B{i}-C{i}-D{i}-E{i}", fmt=CUR)
        cell(ws_mon, i, 7, f"=SUMIFS({ADS('F')},{ADS('B')},$A{i})", fmt=CUR)
        cell(ws_mon, i, 8, f"=F{i}-G{i}", fmt=CUR, bold=True)
        cell(ws_mon, i, 9, f"=IFERROR(H{i}/(B{i}-C{i}),0)", fmt=PCT)
    t = m_last + 1
    cell(ws_mon, t, 1, "Total", bold=True, fill=LIGHT)
    for col in range(2, 9):
        L = chr(64 + col)
        cell(ws_mon, t, col, f"=SUM({L}{m_first}:{L}{m_last})", bold=True, fill=LIGHT, fmt=CUR)
    cell(ws_mon, t, 9, f"=IFERROR(H{t}/(B{t}-C{t}),0)", bold=True, fill=LIGHT, fmt=PCT)
    ws_mon.conditional_formatting.add(
        f"H{m_first}:H{m_last}", CellIsRule(operator="lessThan", formula=["0"], fill=red,
                                            font=Font(name=FONT, color="9C0006")))
    widths(ws_mon, {"A": 12, **{c: 15 for c in "BCDEFGHI"}})
    ws_mon.row_dimensions[4].height = 30

    # ---------------- Summary ----------------
    title(ws_sum, f"{STORE_NAME} - Profit Report", subtitle)
    section(ws_sum, 4, "Key figures")
    kpis = [
        ("Revenue (after discounts)", f"=SUM({ORD('G')})", CUR),
        ("Refunds", f"=SUM({ORD('H')})", CUR),
        ("Net revenue", "=B5-B6", CUR),
        ("Platform fees", f"=SUM({ORD('I')})", CUR),
        ("Cost of goods sold", f"=SUM({ORD('J')})", CUR),
        ("Ad spend", f"=SUM({ADS('F')})", CUR),
        ("Profit after ads", "=B7-B8-B9-B10", CUR),
        ("Profit margin (on net revenue)", "=IFERROR(B11/B7,0)", PCT),
    ]
    for i, (label, formula, fmt) in enumerate(kpis, start=5):
        strong = label.startswith("Profit")
        cell(ws_sum, i, 1, label, bold=strong, fill=LIGHT if strong else None)
        cell(ws_sum, i, 2, formula, fmt=fmt, bold=strong, fill=LIGHT if strong else None)
    n_flag = f"SUMPRODUCT(--(LEN({PROD('J')})>0))"
    cell(ws_sum, 13, 1,
         f'=IF({n_flag}>0,"Warning: "&{n_flag}&" product(s) have no unit cost, so profit is overstated. '
         f'See Data Quality.","")', bold=True, color="C00000", italic=True)

    section(ws_sum, 15, "By channel (ad spend charged to the channel it advertises)")
    for j, h in enumerate(["Channel", "Revenue", "Refunds", "Fees", "COGS", "Ad spend",
                           "Profit after ads", "Margin"], 1):
        header(ws_sum, 16, j, h)
    for i, ch in enumerate(CHANNELS, start=17):
        cell(ws_sum, i, 1, ch)
        for col, src_col in zip(range(2, 6), ["G", "H", "I", "J"]):
            cell(ws_sum, i, col, f"=SUMIFS({ORD(src_col)},{ORD('D')},$A{i})", fmt=CUR)
        cell(ws_sum, i, 6, f"=SUMIFS({ADS('F')},{ADS('E')},$A{i})", fmt=CUR)
        cell(ws_sum, i, 7, f"=B{i}-C{i}-D{i}-E{i}-F{i}", fmt=CUR, bold=True)
        cell(ws_sum, i, 8, f"=IFERROR(G{i}/(B{i}-C{i}),0)", fmt=PCT)
    tr = 17 + len(CHANNELS)
    cell(ws_sum, tr, 1, "Total", bold=True, fill=LIGHT)
    for col in range(2, 8):
        L = chr(64 + col)
        cell(ws_sum, tr, col, f"=SUM({L}17:{L}{tr - 1})", bold=True, fill=LIGHT, fmt=CUR)
    cell(ws_sum, tr, 8, f"=IFERROR(G{tr}/(B{tr}-C{tr}),0)", bold=True, fill=LIGHT, fmt=PCT)

    r0 = tr + 2
    section(ws_sum, r0, "Best and worst products (profit before ads; SKUs with missing cost excluded)")
    for j, h in enumerate(["Rank", "SKU", "Product", "Profit", "Margin"], 1):
        header(ws_sum, r0 + 1, j, h)
    rank_rows = [("Top 1", "LARGE", 1), ("Top 2", "LARGE", 2), ("Top 3", "LARGE", 3),
                 ("Bottom 1", "SMALL", 1), ("Bottom 2", "SMALL", 2), ("Bottom 3", "SMALL", 3)]
    for i, (label, fn, k) in enumerate(rank_rows, start=r0 + 2):
        cell(ws_sum, i, 1, label, bold=True)
        cell(ws_sum, i, 4, f"=IFERROR({fn}({PROD('K')},{k}),0)", fmt=CUR)
        match = f"MATCH($D{i},{PROD('K')},0)"
        cell(ws_sum, i, 2, f'=IFERROR(INDEX({PROD("A")},{match}),"")')
        cell(ws_sum, i, 3, f'=IFERROR(INDEX({PROD("B")},{match}),"")')
        cell(ws_sum, i, 5, f"=IFERROR(INDEX({PROD('I')},{match}),0)", fmt=PCT)

    # charts
    chart_row = r0 + 10
    c1 = BarChart()
    c1.type, c1.grouping = "col", "clustered"
    c1.add_data(Reference(ws_mon, min_col=2, min_row=4, max_row=m_last), titles_from_data=True)
    c1.add_data(Reference(ws_mon, min_col=8, min_row=4, max_row=m_last), titles_from_data=True)
    c1.set_categories(Reference(ws_mon, min_col=1, min_row=m_first, max_row=m_last))
    style_chart(c1, "Revenue and profit after ads, by month", y_title="USD")
    # Explicit series colors and legend labels (blue = revenue, orange = profit)
    for series, label, color in zip(
            c1.series,
            ["Revenue (after discounts)", "Profit (after fees, COGS and ads)"],
            ["2F5597", "C55A11"]):
        series.tx = SeriesLabel(v=label)
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = color
    c1.y_axis.number_format = '$#,##0'
    c1.width, c1.height = 16, 8.5
    ws_sum.add_chart(c1, f"A{chart_row}")

    c2 = BarChart()
    c2.type = "bar"
    c2.add_data(Reference(ws_prod, min_col=11, min_row=4, max_row=p_last), titles_from_data=True)
    c2.set_categories(Reference(ws_prod, min_col=1, min_row=p_first, max_row=p_last))
    style_chart(c2, "Profit by product, before ads", legend=False)
    c2.y_axis.number_format = '$#,##0'
    c2.x_axis.scaling.orientation = "maxMin"
    c2.width, c2.height = 16, 8.5
    ws_sum.add_chart(c2, f"E{chart_row}")

    widths(ws_sum, {"A": 32, "B": 14, "C": 28, "D": 14, "E": 14, "F": 14, "G": 16, "H": 12})

    # ---------------- Data Quality ----------------
    title(ws_dq, "Data Quality Log", "Everything the script found in the raw files, and what it did about it.")
    widths(ws_dq, {"A": 24, "B": 14, "C": 58, "D": 62})
    r = 4
    section(ws_dq, r, "1. Rows read vs. rows used")
    for j, h in enumerate(["File", "Rows read", "Rows used"], 1):
        header(ws_dq, r + 1, j, h)
    for i, (f, n_in, n_out) in enumerate(log.counts, start=r + 2):
        cell(ws_dq, i, 1, f)
        cell(ws_dq, i, 2, n_in, fmt="#,##0")
        cell(ws_dq, i, 3, n_out, fmt="#,##0", align="left")
    r = r + 2 + len(log.counts) + 1

    section(ws_dq, r, "2. Issues found and fixed automatically")
    for j, h in enumerate(["File", "Rows affected", "Issue", "Action taken"], 1):
        header(ws_dq, r + 1, j, h)
    for i, (src, issue, n, action) in enumerate(log.fixes, start=r + 2):
        cell(ws_dq, i, 1, src, wrap=True)
        cell(ws_dq, i, 2, n, fmt="#,##0")
        cell(ws_dq, i, 3, issue, wrap=True)
        cell(ws_dq, i, 4, action, wrap=True)
    r = r + 2 + max(len(log.fixes), 1) + 1

    section(ws_dq, r, "3. Rows needing your attention")
    for j, h in enumerate(["File", "SKU", "Reason", "Order / reference"], 1):
        header(ws_dq, r + 1, j, h)
    if log.attention:
        for i, (src, sku, reason, ref) in enumerate(log.attention, start=r + 2):
            cell(ws_dq, i, 1, src, wrap=True)
            cell(ws_dq, i, 2, sku)
            cell(ws_dq, i, 3, reason, wrap=True)
            cell(ws_dq, i, 4, ref, wrap=True)
    else:
        cell(ws_dq, r + 2, 1, "Nothing to review.")

    # ---------------- Notes ----------------
    title(ws_notes, "Definitions and Assumptions", "Read this before relying on the numbers.")
    notes = [
        ("Sample data", "This workbook was built from randomly generated data for a fictional store. "
                        "No real business figures are shown."),
        ("Revenue", "Shopify: quantity x unit price - discount. Amazon: item-price (already a line total) "
                    "- promotion discount. Shipping charged to customers is not included."),
        ("Refunds", "Shopify refunds come from the Refunded Amount column. Amazon refunds arrive as "
                    "separate negative rows and are attached to the original order, so a refund counts "
                    "in the month of the original sale."),
        ("Fees", f"Shopify: {SHOPIFY_FEE_PCT:.1%} of line revenue + ${SHOPIFY_FEE_FIXED:.2f} per order "
                 "(an assumption - replace with the client's real payment rates). Amazon: the amazon-fees "
                 "column as reported. Fees are not refunded when an order is refunded."),
        ("COGS", "Units x unit cost from the cost file. Cost of refunded units is NOT recovered "
                 "(a conservative assumption). SKUs missing from the cost file get $0 COGS and are flagged."),
        ("Ad spend", "Meta Ads and Google Ads are charged to Shopify; Amazon Ads to Amazon (see "
                     "AD_PLATFORM_TO_CHANNEL in the script). Ad spend is not allocated to individual "
                     "products, so product profit is shown before ads."),
        ("Dates", "Shopify timestamps are used as exported (store time zone). Amazon dates are parsed as "
                  "MM/DD/YYYY."),
        ("Profit", "Net revenue - fees - COGS - ad spend. Margin = profit / net revenue."),
        ("Formulas", "All figures on Summary, By Product and By Month are formulas reading from the "
                     "Orders and Ad Spend sheets. The yellow cell on By Product (margin flag) is an input."),
    ]
    widths(ws_notes, {"A": 18, "B": 110})
    for i, (k, v) in enumerate(notes, start=4):
        cell(ws_notes, i, 1, k, bold=True)
        cell(ws_notes, i, 2, v, wrap=True)
        ws_notes.row_dimensions[i].height = 32

    for ws in (ws_sum, ws_prod, ws_mon, ws_dq, ws_notes):
        ws.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)


# =====================================================================================
# Main
# =====================================================================================
def main():
    ap = argparse.ArgumentParser(description="Build a multi-channel profit report from messy exports.")
    ap.add_argument("--input-dir", default="sample_data", help="Folder containing the four input files")
    ap.add_argument("--output", default="northwind_profit_report.xlsx", help="Output .xlsx path")
    args = ap.parse_args()

    t0 = time.time()
    folder, log = Path(args.input_dir), QualityLog()

    shopify = load_shopify(folder / "shopify_orders.csv", log)
    amazon = load_amazon(folder / "amazon_report.csv", log)
    ads = load_ads(folder / "ad_spend.csv", log)
    costs = load_costs(folder / "product_costs.xlsx", log)
    orders = combine(shopify, amazon, costs, log)
    write_workbook(orders, ads, log, Path(args.output))

    rows_in = sum(c[1] for c in log.counts)
    print(f"Read {rows_in:,} raw rows from 4 files")
    print(f"Fixed {sum(f[2] for f in log.fixes):,} data issues across {len(log.fixes)} categories")
    print(f"Wrote {len(orders):,} clean order lines and {len(ads):,} ad rows")
    if log.attention:
        print(f"{len(log.attention)} item(s) need review - see the 'Data Quality' sheet")
    print(f"Saved {args.output} in {time.time() - t0:.1f}s")


if __name__ == "__main__":
    main()
