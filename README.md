# Financial Controlling Dashboard — Excel Power Pivot

A navigable Excel controlbook for project portfolio management, built with Power Query, Power Pivot, DAX, and CUBE formulas. Designed for engineering & construction companies tracking project financials across multiple regions.

![Excel](https://img.shields.io/badge/Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Power Pivot](https://img.shields.io/badge/Power%20Pivot-DAX-orange)
![CUBE Functions](https://img.shields.io/badge/CUBE-Functions-blueviolet)
![License](https://img.shields.io/badge/License-MIT-blue)

## Overview

This workbook demonstrates a **navigable controlbook** approach — a single Excel file with interconnected sheets for drilling down from summary to regional detail views.

**Use case:** Portfolio Controller at a multi-regional engineering company needs to monitor project health, cash flow, and margin performance across 66 active projects.

## Features

- **Star Schema Data Model** — 6 dimension tables + 6 fact tables (1,344 rows)
- **DAX Measures** — Margin, Unbilled, Cash Flow, AR, Customer Advances (all USD-converted)
- **CUBE Formulas** — `CUBEVALUE`, `CUBERANKEDMEMBER`, `CUBESET` for flexible cell-based reporting
- **PivotTables** — Regional breakdowns with calculated measures
- **Interactive Slicers** — Filter by Region, Project Status, Time Period
- **Conditional Formatting** — Top/Bottom performers highlighted automatically
- **Data Dictionary** — Built-in KPI definitions and sanity checks

## Technical Approach

This workbook uses **two complementary techniques**:

| Technique | Used For | Sheets |
|-----------|----------|--------|
| **CUBE Formulas** | KPI cards, dynamic lookups, ranked lists | Summary, Support |
| **PivotTables** | Regional breakdowns, project lists, cash analysis | Summary, Subsidiary Dashboard, Cash Dashboard |

### Why Both?

- **CUBE formulas** give precise cell-level control — perfect for KPI strips and custom layouts
- **PivotTables** handle large data aggregations efficiently with built-in slicer support
- Combined approach = flexibility of formulas + power of pivot engine

## Workbook Structure

```
Navigating_Files_Project_Controlbook_v1.xlsx
│
├── Summary                 ← Portfolio KPIs (CUBE + PivotTable)
├── Subsidiary Dashboard    ← Regional drill-down (PivotTable)
├── Cash Dashboard          ← AR, Unbilled, Cash Flow (PivotTable + CUBE)
├── DATA DICTIONARY         ← KPI definitions & sanity checks
├── Support                 ← Helper calculations (CUBESET)
│
└── [Data Model - hidden]
    ├── dim_client
    ├── dim_company
    ├── dim_date_month
    ├── dim_project
    ├── dim_region
    ├── dim_vendor
    ├── fact_project_month      ← main fact (1,344 rows)
    ├── fact_project_contract
    ├── fact_baseline_month
    ├── fact_resource_plan_month
    ├── fact_commitments
    └── fact_forecast_snapshot
```

## CUBE Formulas Used

```excel
// Get KPI value for entire portfolio
=CUBEVALUE("ThisWorkbookDataModel"; "[Measures].[Margin]")

// Get top 3 projects by Earned Revenue
=CUBERANKEDMEMBER("ThisWorkbookDataModel"; 
    "[dim_project].[project_name_pretty]"; 1;
    "[Measures].[Earned]")

// Create a set for filtering
=CUBESET("ThisWorkbookDataModel"; 
    "[dim_region].[region_name].&[West]"; "West Region")
```

## Key Metrics

| KPI | Description | Good When |
|-----|-------------|-----------|
| EARNED | Revenue recognized (% complete × contract) | Higher |
| COST | Total project costs incurred | Lower |
| MARGIN | Earned − Cost | Higher / Positive |
| MARGIN % | Margin ÷ Earned | Higher |
| UNBILLED | Earned − Invoiced (negative = overbilled) | Context-dependent |
| CASH IN | Cash received from clients | Higher |
| CASH OUT | Cash paid (materials, subcontractors) | Lower |
| CASH FLOW | Cash In − Cash Out | Positive |
| AR | Receivables (Invoiced − Cash Received) | Lower |
| CUSTOMER ADVANCES | Prepayments received before invoicing | Context-dependent |

## Data Model

```
                    ┌─────────────┐
                    │  dim_date   │
                    │   _month    │
                    └──────┬──────┘
                           │
┌───────────┐    ┌─────────┴─────────┐    ┌───────────┐
│dim_client │────│ fact_project_month│────│dim_vendor │
└───────────┘    │    (main fact)    │    └───────────┘
                 └─────────┬─────────┘
                           │
              ┌────────────┼────────────┐
              │            │            │
       ┌──────┴──────┐ ┌───┴────┐ ┌─────┴─────┐
       │ dim_project │ │dim_    │ │dim_company│
       │             │ │region  │ │           │
       └─────────────┘ └────────┘ └───────────┘
```

**Additional fact tables:** fact_baseline_month, fact_commitments, fact_forecast_snapshot, fact_project_contract, fact_resource_plan_month

## Sample DAX Measures

```dax
// Margin (USD) - converts from PLN using fx_rate
Margin := 
SUMX(
    fact_project_month;
    DIVIDE(
        [earned_revenue_pln] - [total_cost_pln];
        [fx_rate]
    )
)

// AR (Receivables) - only positive values
AR := 
SUMX(
    fact_project_month;
    VAR ar_value = DIVIDE([invoiced_revenue_pln] - [cash_received_pln]; [fx_rate])
    RETURN IF(ar_value > 0; ar_value; 0)
)

// Customer Advances - prepayments (cash > invoiced)
Customer Advances := 
SUMX(
    fact_project_month;
    VAR adv_value = DIVIDE([cash_received_pln] - [invoiced_revenue_pln]; [fx_rate])
    RETURN IF(adv_value > 0; adv_value; 0)
)
```

> **Note:** DAX uses semicolon (`;`) as separator in European locale Excel.

## File Structure

```
Financial-Controlbook-Excel/
├── README.md
├── Navigating_Files_Project_Controlbook_v1.xlsx
├── data/
│   ├── dim_client.csv
│   ├── dim_company.csv
│   ├── dim_date_month.csv
│   ├── dim_project.csv
│   ├── dim_region.csv
│   ├── dim_vendor.csv
│   ├── fact_baseline_month.csv
│   ├── fact_commitments.csv
│   ├── fact_forecast_snapshot.csv
│   ├── fact_project_contract.csv
│   ├── fact_project_month.csv
│   └── fact_resource_plan_month.csv
└── screenshots/
    └── (add your own)
```

## Getting Started

1. Download `Navigating_Files_Project_Controlbook_v1.xlsx`
2. Enable content if prompted (for Power Pivot)
3. Start on **Summary** sheet
4. Use slicers to filter by Region or Status
5. Check **DATA DICTIONARY** for KPI explanations

### Requirements

- Microsoft Excel 2016+ (Windows) or Excel 365
- Power Pivot add-in enabled
- ~50 MB RAM for data model

## Portfolio Summary

| Metric | Value |
|--------|-------|
| Total Projects | 66 |
| Regions | 4 (West, Midwest, South, Northeast) |
| BAC (Budget at Completion) | $916M |
| EAC (Estimate at Completion) | $821M |
| Total Commitments | $9.9B |
| Data Rows | 1,344 |

## Color Scheme (Northbridge Brand)

| Element | Color | Hex |
|---------|-------|-----|
| Header Navy | Dark Blue | `#0B1F33` |
| Lines/Borders | Light Gray | `#CBD5E1` |
| Text | Gray | `#6B7280` |
| Positive/Green | Status OK | `#166534` |
| Warning/Amber | Attention | `#B45309` |
| Alert/Red | Problem | `#991B1B` |

## Lessons Learned

- **CUBE + PivotTables** — Use both: CUBE for precise control, PivotTables for aggregations
- **Single slicer per field** — Multiple slicers for same field cause sync issues
- **VBA limitations** — Macro security often blocks custom highlighting; use Conditional Formatting instead
- **Negative values need context** — Unbilled < 0 means overbilled (advance billing), not an error
- **Data Dictionary is essential** — Document KPI logic to avoid misinterpretation
- **DAX locale matters** — Use semicolon (`;`) in European Excel, comma (`,`) in US

## Skills Demonstrated

- Excel Data Modeling (Star Schema)
- Power Query (ETL)
- Power Pivot & DAX
- CUBE Functions (`CUBEVALUE`, `CUBERANKEDMEMBER`, `CUBESET`)
- PivotTables & Slicers
- Conditional Formatting
- Dashboard Design
- Financial KPI Logic

## License

MIT — feel free to use as a template for your own projects.

## Author

Built as a portfolio project demonstrating Excel data modeling and dashboard design skills.

---

*Synthetic data only — Northbridge Engineering & Construction Ltd. is a fictional company.*
