# From four messy exports to one trustworthy profit report

*A portfolio project built on randomly generated data for a fictional store. No real business figures are shown.*

**The problem.** Northwind Candle Co. (fictional) sells on Shopify and Amazon and advertises on Meta, Google and Amazon Ads. Every month the owner downloads four files that don't agree with each other: different column names, different date formats, refunds recorded in different ways, duplicate rows, and SKUs typed five different ways. Working out which products and channels actually make money means a long session of manual copy, paste and VLOOKUP.

**What I built.** A Python script that takes the four raw exports and produces one formatted Excel workbook:

- **Summary:** revenue, refunds, fees, cost of goods, ad spend and profit, plus a channel breakdown, best and worst products, and two charts.
- **By Product and By Month:** profit and margin per SKU, and the monthly trend.
- **Data Quality:** a plain-English log of everything the script found and fixed, with a short list of rows that need a human to look at them.
- **Live formulas:** every number on the report sheets is an Excel formula over the clean data, so the workbook still works if someone edits a row.

**What the script cleans.**

- Duplicate and blank rows
- SKUs with stray spaces, mixed case and underscores, so they match across all four files
- Prices and costs stored as text with $ symbols
- Mismatched date formats (parsed explicitly, so day and month can never be swapped)
- Amazon refunds that arrive as separate negative rows, attached back to the original order
- Different price definitions (Shopify lists a unit price, Amazon lists a line total)

**The design choice I'm proudest of.** Nothing is dropped or guessed silently. In the sample data, one product sells but has no cost on file. Rather than hiding it or inventing a number, the script sets its cost to zero, highlights it, ranks it last, excludes it from the best and worst lists, and puts a warning at the top of the Summary. A report that admits what it doesn't know is one a business owner can trust.

**Results on the sample data.**

| Measure | Result |
|---|---|
| Raw rows read | 3,503 across 4 files |
| Data issues fixed automatically | 961 across 14 categories |
| Clean order lines produced | 2,321 |
| Formulas in the finished workbook | 2,518, with zero errors |
| Run time | About 1.5 seconds |

I also checked the totals by recomputing revenue, refunds, fees and ad spend from the raw files using separate logic. Everything matched to the cent.

**What the report revealed.** In the sample data, Shopify earned a 39.9% margin after ads while Amazon earned only 6.9%. One product, wax melts, had the weakest margin at 18.5%. That is the kind of insight that stays hidden when profit is calculated by hand across four files.

**Limitations I documented.** Ad spend is attributed by platform, not by product. Shopify fees are an explicit assumption. Refunded units are treated as not restocked. These are listed on the report's Notes sheet and in the [README](README.md), because a client should know what a number means before relying on it.

**Tools.** Python, pandas, openpyxl, Excel formulas (SUMIFS, INDEX/MATCH), conditional formatting, native Excel charts.

**Links.** Code and sample data: this repository. Example workbook: [northwind_profit_report.xlsx](northwind_profit_report.xlsx).

## Contact

Want this for your own store or business? Charles Maxwell, Toronto, Ontario
[charlesmaxwell87@gmail.com](mailto:charlesmaxwell87@gmail.com) | [LinkedIn](https://www.linkedin.com/in/charles-maxwell-93a467232/)
