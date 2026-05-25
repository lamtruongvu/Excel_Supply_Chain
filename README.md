# Supply Chain Analytics — Excel Project

A formula-driven Excel workbook that transforms a 100-SKU, 24-variable supply chain dataset into a fully automated analytics suite. Built to demonstrate production-level Excel competency across data modelling, multi-criteria aggregation, KPI reporting, and stakeholder-ready templates.

---

## Project Overview

| Attribute | Detail |
|-----------|--------|
| **Dataset** | 100 SKUs × 24 variables (product, supplier, logistics, inventory, quality, sales) |
| **Workbook sheets** | 8 (1 raw data, 5 analysis, 2 templates) |
| **Live formulas** | 337 |
| **Formula errors** | 0 |
| **Tools** | Microsoft Excel, Python (openpyxl), pandas |

---

## Files

```
supply_chain_data.csv                  — Source dataset (100 rows × 24 columns)
Supply_Chain_Formula_Workbook.xlsx     — Main deliverable: formula-driven workbook
Supply_Chain_Report.docx               — Full project report (methodology, findings, formula catalogue)
```

---

## Workbook Structure

### Raw Data
- Full dataset loaded as a named Excel Table (`RawData`)
- Structured column references (e.g. `RawData[Product type]`) allow all downstream formulas to auto-expand if rows are added
- Alternating row colours, frozen panes, and number formatting applied across all 24 columns

### Analysis Sheets (5)

#### 1. Revenue by Product
Pivots by product type (Cosmetics / Haircare / Skincare) across revenue, units sold, pricing, defect rate, and lead time.

Key formulas:
```excel
=SUMIF(RawData[Product type], A5, RawData[Revenue generated])
=AVERAGEIF(RawData[Product type], A5, RawData[Defect rates]) / 100
=D5 / SUM($D$5:$D$7)   ← Revenue share with locked denominator
```

#### 2. Supplier Performance
Ranks all 5 suppliers across lead time, manufacturing cost, defect rate, production volume, and inspection pass/fail rates.

Key formulas:
```excel
=AVERAGEIF(RawData[Supplier name], A5, RawData[Manufacturing costs])
=COUNTIFS(RawData[Supplier name], A5, RawData[Inspection results], "Pass") / COUNTIF(RawData[Supplier name], A5)
```
> `COUNTIFS` is required here because splitting inspection results by outcome (Pass/Fail/Pending) cannot be done with a native pivot calculated field — it needs two simultaneous criteria.

#### 3. Logistics & Shipping
Three sub-tables in one sheet: by shipping carrier, by transportation mode, and by route. Uses both `Shipping costs` (carrier fee) and `Costs` (total logistics cost) intentionally — the distinction matters for cost analysis accuracy.

Key formulas:
```excel
=SUMIF(RawData[Shipping carriers], A6, RawData[Shipping costs])
=AVERAGEIF(RawData[Transportation modes], G6, RawData[Costs])
=SUMIF(RawData[Routes], A14, RawData[Costs]) / D14   ← Cost-to-Revenue ratio
```

#### 4. Location & Inventory
Groups by warehouse city (Bangalore, Chennai, Delhi, Kolkata, Mumbai) with two derived metrics not available as native aggregations.

Key formulas:
```excel
=AVERAGEIF(RawData[Location], A5, RawData[Availability]) / 100
=IFERROR(C5 / G5, "-")   ← Stock-to-Order ratio
=IFERROR(F5 / B5, "-")   ← Revenue per SKU
```

#### 5. Quality & Defects
Two linked tables (by product type and by supplier) with full Pass/Fail/Pending breakdowns and derived rate columns.

Key formulas:
```excel
=COUNTIFS(RawData[Product type], A6, RawData[Inspection results], "Pass")
=COUNTIFS(RawData[Product type], A6, RawData[Inspection results], "Fail")
=IFERROR(C6 / F6, 0)   ← Pass Rate %
```

### Reporting Templates (2)

#### KPI Report Template
- 4 live headline KPI cards pulling from analysis sheet totals rows via cross-sheet references
- 14-row KPI table with current value (live), target, and traffic-light status
- Pre-written findings and recommendations based on actual data

Cross-sheet reference examples:
```excel
='Revenue by Product'!D8        ← Total Revenue
='Quality & Defects'!G9         ← Overall Pass Rate
='Logistics & Shipping'!C9      ← Avg Shipping Time
```

#### Monthly Report Template
- 8-section structured template for recurring reporting cycles
- Metric tables with pre-filled variance formulas: `=IFERROR(C-B, "")`

---

## Key Findings

| Area | Finding | Priority |
|------|---------|----------|
| Quality | Supplier 4: **0 passing inspections** out of 18 SKUs | 🔴 Critical |
| Quality | Overall pass rate **23%** vs 40% target; 36% actively failing | 🔴 Critical |
| Quality | Supplier 5 carries the highest defect rate at **2.67%** | 🔴 High |
| Inventory | Avg availability **48.4%** vs 60% target | 🟡 Monitor |
| Logistics | Supplier 3 lead time **20.1 days** vs 15-day target | 🟡 Monitor |
| Revenue | Skincare drives **41.8% of total revenue** ($241,628) | 🟢 Positive |
| Logistics | Carrier B: most shipments, fastest transit, competitive cost | 🟢 Positive |
| Supplier | Supplier 1 benchmarks best across all quality and cost dimensions | 🟢 Positive |

---

## Formulas Used

| Function | Purpose |
|----------|---------|
| `SUMIF` | Sum a column for rows matching one criterion |
| `AVERAGEIF` | Average a column for rows matching one criterion |
| `COUNTIF` | Count rows matching one criterion |
| `COUNTIFS` | Count rows matching **two** criteria simultaneously |
| `SUM` / `AVERAGE` | Totals rows aggregation |
| `IFERROR` | Suppress division errors in derived ratio columns |
| Cross-sheet `=` | Pull live values into KPI dashboard from analysis sheets |
| Locked `$` reference | Fix denominator in revenue share formula when copying |

---

## Conditional Formatting

| Format | Applied to | Meaning |
|--------|-----------|---------|
| 3-colour scale (Green → Red) | Defect rate columns | Higher defect = more red |
| 3-colour scale (Red → Green) | Pass rate / Availability | Lower value = more red |
| Data bars (blue) | Production volumes | Visual output proportion |
| Data bars (teal) | Revenue by location | Relative revenue contribution |

---

## Skills Demonstrated

- Advanced Excel: SUMIF, AVERAGEIF, COUNTIFS, IFERROR, cross-sheet references, locked references
- Data modelling: named Excel Tables, structured column references, auto-expanding formula ranges
- KPI framework design: derived metrics (pass rate, revenue share, stock-to-order ratio, cost-to-revenue ratio)
- Conditional formatting: colour scales and data bars for visual performance signalling
- Dashboard design: live KPI cards, traffic-light status indicators, chart embedding
- Reporting templates: reusable monthly report structure with pre-built variance formulas
- Python + openpyxl: programmatic workbook generation and formula validation

---

## How to Use

1. Open `Supply_Chain_Formula_Workbook.xlsx`
2. Navigate to **Raw Data** — this is the single source of truth
3. All five analysis sheets auto-calculate from this table; no manual refresh required
4. To update with new data: paste new rows into the Raw Data table — all formulas, charts, and KPI cards update automatically
5. Use **KPI Report Template** for executive summaries; use **Monthly Report Template** for recurring operational reviews

---

## Author

**Lam Truong Vu**
Master of Information Technology — Data Analytics | Kaplan Business School
[LinkedIn](https://www.linkedin.com/in/larry-vu9/) | vulamtruong9501@gmail.com
