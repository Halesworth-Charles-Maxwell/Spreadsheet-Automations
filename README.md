# Multi-Channel Profit Report Automation

Turns four messy monthly exports into one clean, formula-driven Excel profit report in about two seconds.

**The scenario (fictional):** Northwind Candle Co. sells on Shopify and Amazon and advertises on Meta, Google and Amazon Ads. Each month the owner downloads four files and spends hours in Excel trying to work out what is actually profitable. This script replaces that process.

> All data in this project is randomly generated. No real business figures are used.

## Quick start

```bash
pip install pandas openpyxl

python generate_sample_data.py      # creates sample_data/ (4 messy files)
python build_report.py              # creates northwind_profit_report.xlsx
```

Options: `python build_report.py --input-dir my_files --output report.xlsx`

## What goes in

| File | What is wrong with it (on purpose) |
|---|---|
| `shopify_orders.csv` | Duplicate rows, blank rows, rows with no SKU, SKUs like `cnd_03 ` / ` CND-03`, prices stored as `$24.00` text, timestamps with time zones, refunds in a column |
| `amazon_report.csv` | Different column names, `MM/DD/YYYY` dates, item price is a **line total** (Shopify's is per unit), refunds as separate **negative rows**, `$` in numbers, duplicates |
| `ad_spend.csv` | Platform names in three capitalizations, `$` in spend values |
| `product_costs.xlsx` | Messy SKUs, one cost stored as text, and one SKU that sells but has **no cost entry** |

## What comes out

`northwind_profit_report.xlsx`:

| Sheet | Contents |
|---|---|
| **Summary** | Key figures, profit by channel, best/worst products, two charts |
| **By Product** | Profit and margin per SKU (before ad spend), low margins highlighted |
| **By Month** | Revenue, costs, ad spend and profit per month |
| **Data Quality** | Rows read vs. used, every issue fixed, and rows that need a human to look at them |
| **Orders / Ad Spend** | The clean merged data that all report formulas read from |
| **Notes** | Definitions and assumptions |

Every figure on the report sheets is an Excel formula (`SUMIFS` etc.) over the clean data sheets, so the workbook keeps working if someone edits or adds a row.

## Design decisions worth knowing

- **Nothing is dropped silently.** Rows that cannot be used (e.g. no SKU) are removed *and* listed on the Data Quality sheet with their order number.
- **Missing costs are flagged, not guessed.** A SKU with no unit cost gets $0 COGS, is highlighted, sorts to the bottom of the product table, is excluded from the best/worst rankings, and triggers a warning on the Summary.
- **Dates are parsed with explicit formats**, so `04/03/2026` can never be misread as 4 March.
- **Refunds count in the month of the original sale.**
- **Ad spend is not split by product**, because the ad data does not say which product an ad promoted. Product profit is therefore "before ads"; channel and month profit are "after ads".
- **Fee assumptions are explicit.** Shopify fees (2.9% + $0.30 per order) are an assumption set at the top of `build_report.py`; Amazon fees come from the report.

## Adapting this to a real client

1. Put the client's exports in a folder and point `--input-dir` at it.
2. In `build_report.py`, adjust the column names inside `load_shopify`, `load_amazon`, `load_ads` and `load_costs` to match the client's files.
3. Update the settings block at the top: store name, fee rates, the ad-platform-to-channel mapping, margin threshold.
4. Run it, then read the **Data Quality** sheet first. It tells you what to ask the client about.

## Known limitations

- Formulas calculate when the file is opened in Excel (or LibreOffice). Some previewers show blanks until then.
- Ad spend is attributed by platform-to-channel mapping only, not by campaign-level sales attribution.
- Cost of refunded units is treated as not recovered. Change this in `combine()` if the client restocks returns.
- Shipping charged to customers and shipping costs are not modeled.

## Files

```
generate_sample_data.py   Creates the messy sample inputs (seeded, repeatable)
build_report.py           Cleans, merges and writes the Excel report
sample_data/              The four generated input files
northwind_profit_report.xlsx   Example output
```
