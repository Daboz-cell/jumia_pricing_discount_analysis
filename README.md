# Jumia Home & Kitchen Category — Pricing, Discount & Rating Analysis

An end-to-end Excel data analytics project that cleans a raw Jumia (Home & Kitchen category) product export, engineers pricing/rating features, builds KPI pivot tables, and packages the results into an interactive dashboard — aimed at answering one core business question for Jumia and its sellers:

> **Does discounting actually drive engagement, and how should sellers price and manage their catalog as a result?**

---

## 📁 Repository Contents

| File | Description |
|---|---|
| `jumia_clean_dataset.xlsx` | The full workbook — raw data, cleaned data, feature-engineered columns, KPI summary pivot tables, and the interactive Dashboard sheet |
| `README.md` | This file — process, findings, and recommendations |

---

## 1. Project Scope & Business Intent

The dashboard was built to support three decisions:

1. **Promotional Strategy & Margin Optimization** — do deep discounts genuinely drive volume (reviews), or do they just shrink margins?
2. **Listing Quality & Seller Compliance** — flag unrated or low-rated (<3.0) products for quality control, re-listing, or review campaigns.
3. **Catalog Pricing & Segmentation** — price products into clear market bands (Budget / Mid-Range / Premium) to maximize market share.

---

## 2. Data Cleaning & Preprocessing (Excel)

Raw data (115 rows) was copied into a `Cleaned Data` sheet, leaving the original sheet untouched as a backup.

| Issue | Fix Applied |
|---|---|
| Exact duplicate rows (43, 84, 96) | `Data → Remove Duplicates` → 115 rows → **112 rows** |
| `"KSh 1,525"` price strings | `Find & Replace` / `=VALUE(SUBSTITUTE(SUBSTITUTE(B2,"KSh ",""),",",""))` |
| Price ranges e.g. `"KSh 1,620 - KSh 1,980"` | Averaged lower/upper bound manually for the single affected row (Row 39, sofa cover) → Current Price midpoint **1,800**, Old Price midpoint **2,700** |
| `Discount` stored as text with `%` | Stripped `%`, divided by 100 via Paste Special → Divide, formatted as Percentage |
| Negative review counts | `Paste Special → Multiply by −1` (equivalent to `ABS()`) |
| Missing review counts (blanks) | `Go To Special → Blanks` → filled with `0` |
| `Rating` stored as `"4.5 out of 5"` text | `Find & Replace` removed `" out of 5"`, parsed to numeric |
| Missing ratings (55 blank rows) | Filled with `"Not Rated"` label, excluded from numeric rating averages |
| Inconsistent headers (`old price`, `Ratingd`) | Standardized to `Old_Price`, `Rating` |

---

## 3. Feature Engineering (Cleaned Data sheet, Columns G–J)

| Column | Formula | Logic |
|---|---|---|
| **Discount Amount** (G) | `=C2-B2` | Old Price − Current Price |
| **Price Category** (H) | `=IF(B2<=500,"Budget",IF(B2<=1500,"Mid-Range","Premium"))` | Budget ≤500, Mid-Range 501–1,500, Premium >1,500 |
| **Discount Category** (I) | `=IF(D2<0.2,"Low (<20%)",IF(D2<=0.4,"Medium (20%-40%)","High (>40%)"))` | Low / Medium / High discount depth |
| **Rating Category** (J) | `=IFS(F2="Not Rated","Not Rated",F2<3,"Poor",F2<=4,"Average",F2<=4.5,"Good",F2>4.5,"Excellent")` | Closes the 4.0–4.5 gap in the original brief with a continuous nested `IFS` |

---

## 4. Baseline KPIs

| Metric | Formula | Result |
|---|---|---|
| Total Cleaned Products | `=COUNTA(A2:A113)` | **112** |
| Total Reviews | `=SUM(E2:E113)` | **723** |
| Average Current Price | `=AVERAGE(B2:B113)` | **KSh 1,186.89** |
| Average Old Price | `=AVERAGE(C2:C113)` | **KSh 1,811.11** |
| Average Discount % | `=AVERAGE(D2:D113)` | **36.78%** |
| Average Rating (rated items only) | `=AVERAGEIF(F2:F113,"<>Not Rated")` | **3.89 / 5.0** |
| Most Expensive Product | `=MAX(B2:B113)` | 32PCS Portable Cordless Drill Set — KSh 3,750.00 |
| Least Expensive Product | `=MIN(B2:B113)` | 3PCS Single Head Knitting Crochet — KSh 38.00 |

---

## 5. Pivot Table Analysis

**Pivot 1 — Price Category Distribution**

| Category | Products | Reviews | Avg. Price |
|---|---|---|---|
| Budget (≤KSh 500) | 29 | 232 | KSh 292.14 |
| Mid-Range (KSh 501–1,500) | 46 | 209 | KSh 980.57 |
| Premium (>KSh 1,500) | 37 | 282 | KSh 2,144.70 |

Mid-Range dominates inventory count, but Budget and Premium items pull in more reviews per listing — engagement is strongest at the price extremes, not the middle.

**Pivot 2 — Discount Depth vs. Review Engagement**

| Discount Tier | Products | Reviews |
|---|---|---|
| Low (<20%) | 19 | 50 |
| Medium (20–40%) | 31 | 339 |
| High (>40%) | 62 | 334 |

Medium-to-high discounted listings (≥20%) account for **93% of all reviews (673 of 723)** — a strong signal that discounting drives engagement.

**Pivot 3 — Rating Tier & Catalog Health**

| Rating Tier | Products | Reviews |
|---|---|---|
| Not Rated | 55 | 0 |
| Excellent (>4.5) | 19 | 289 |
| Good (4.1–4.5) | 13 | 172 |
| Average (3.0–4.0) | 13 | 93 |
| Poor (<3.0) | 12 | 169 |

**49% of the catalog (55/112 listings) has zero ratings** — a large pool of "cold" listings needing review-generation campaigns. 12 Poor-rated products still pulled 169 reviews, flagging specific high-traffic items for urgent quality review.

---

## 6. Dashboard

The `Dashboard` sheet presents:
- 4 KPI cards (Total Listings, Avg. Current Price, Avg. Discount Rate, Total Reviews) pinned across the top for instant executive scanning
- Gridlines removed for a clean, presentation-ready canvas
- 3 linked Pivot Charts (Price Category, Discount Depth, Rating Tier)
- Interactive **Slicers** for Price Category and Discount Category, connected via *Report Connections* to all three pivot tables/charts simultaneously — so filtering "Premium" or "High Discount" instantly updates the whole dashboard

---

## 7. Key Findings & Business Insights

**Are higher discounts leading to higher customer engagement?**
Yes. Listings discounted ≥20% generate 93% of all reviews (673/723), while the Low-discount tier (19 products) produced only 50 reviews. Discounting is clearly correlated with engagement in this category.

**Do highly rated products have higher or lower prices?**
Essentially neither — price and rating show almost no relationship (correlation ≈ **+0.11** across the 57 rated products). Customers rate products on quality/fit for purpose, not on how expensive they are. This means Jumia sellers can't assume a higher price signals — or requires — a higher rating; pricing and quality perception should be managed as separate levers.

**Which products are performing best, and which need a better strategy?**
- **Best performers:** low/mid-priced items with 5.0 ratings and steady review counts (e.g. the Aluminum Folding Truck Hand Cart at KSh 2,025 with a 5.0 rating and 3 reviews, several sub-KSh 350 home-decor pieces also rated 5.0). These show that consistent quality at accessible price points converts well.
- **Needs attention:** the 120W Cordless Vacuum Cleaner (KSh 445, 2.8 rating, 69 reviews) and the VIC Wireless Vacuum Cleaner (KSh 1,220, 2.9 rating, 16 reviews) are high-review but Poor-rated — these are visible, high-traffic listings actively damaging category trust and should be prioritized for vendor quality review or delisting.

---

## 8. Recommendations for Jumia Sellers

1. **Maintain discounts in the 20–40% "Medium" band as the default promotional strategy.** It already captures the large majority of review volume (Medium alone: 31 products → 339 reviews) without the margin erosion of blanket 40%+ discounting, which — while it drives volume (62 products, 334 reviews) — spreads engagement across more listings per review earned.
2. **Launch a targeted review-generation campaign for the 55 unrated listings (49% of the catalog).** These products are invisible in trust signals to buyers; even a small nudge (post-purchase review prompts, small review-linked incentives) would materially close the "Not Rated" gap.
3. **Audit and remediate the 12 Poor-rated (<3.0) products that already carry meaningful review volume (169 reviews total)**, starting with the two vacuum cleaner listings above. These are the highest-risk items because they combine visibility (high reviews) with poor quality perception (low rating) — the worst possible combination for category reputation.

---

## 9. Tools Used

- Microsoft Excel (formulas, Conditional Formatting, Remove Duplicates, Go To Special, Pivot Tables, Pivot Charts, Slicers, Report Connections)
- No external libraries — entirely native Excel workflow, chosen to match the brief's requirement for a reproducible, formula-driven analysis

---

