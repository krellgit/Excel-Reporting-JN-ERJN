# Excel Reporting -JN (ERJN)

Campaign Performance Report Generator - Excel Version

A comprehensive Excel-based reporting solution that replicates the functionality of the HTML Campaign Report Generator with native Excel controls, formulas, and formatting.

## Features

### Interactive Controls
1. **Portfolio Toggle**: Switch between Overall, JN, and Non-JN portfolio views
2. **Time Period Toggle**: Aggregate data by Daily, Weekly, or Monthly periods
3. **Date Range Selection**: Filter data by custom date ranges

### Report Sheets

| Sheet | Description |
|-------|-------------|
| **Instructions** | Getting started guide and metric definitions |
| **Dashboard** | Main control panel with KPIs and summary metrics |
| **Executive Summary** | Key metrics with month-over-month comparisons |
| **Segment Performance** | Branded/Competitor/Non-Branded breakdown |
| **Performance Trends** | Daily/Weekly/Monthly trend analysis |
| **Monthly Analysis** | Detailed month-over-month data |
| **Weekly Analysis** | Detailed week-over-week data |
| **Organic vs Paid** | Total Sales breakdown (requires Business Report) |
| **JN Portfolio** | JN-specific metrics and segments |
| **Non-JN Portfolio** | Non-JN-specific metrics and segments |
| **Pivot - Portfolio** | Cross-tabulation by portfolio and time |
| **Pivot - Segment** | Cross-tabulation by segment and time |
| **Campaign Data** | Raw campaign data input |
| **Business Data** | Raw business report input |
| **Settings** | Configuration and dropdown values |

### Key Metrics Calculated
1. **Ad Spend** - Total advertising spend
2. **Ad Sales** - Sales attributed to advertising (7-day attribution)
3. **ROAS** - Return on Ad Spend = Ad Sales / Ad Spend
4. **ACoS** - Advertising Cost of Sale = Ad Spend / Ad Sales × 100%
5. **TACOS** - Total Advertising Cost of Sale = Ad Spend / Total Sales × 100%
6. **CVR** - Conversion Rate = Orders / Clicks × 100%
7. **CPC** - Cost Per Click = Ad Spend / Clicks
8. **CTR** - Click-Through Rate = Clicks / Impressions × 100%
9. **Organic Sales** - Total Sales - Ad Sales

## Usage

### Option 1: Use Pre-built Template
1. Open `Campaign_Report_Template.xlsx`
2. Paste your Sponsored Products Campaign Report CSV data into the 'Campaign Data' sheet (starting row 5)
3. Paste your Business Report CSV data into the 'Business Data' sheet (starting row 5)
4. Adjust date range in 'Settings' sheet (cells B6 and B7)
5. Use Portfolio and Time Period dropdowns in 'Settings' to filter views

### Option 2: Generate Fresh Template
```bash
pip install openpyxl
python generate_excel_report.py
```

## Data Input Requirements

### Campaign Data (Required)
Export your Sponsored Products Campaign Report from Amazon Advertising with these columns:
1. Date
2. Portfolio name
3. Campaign Name
4. Impressions
5. Clicks
6. Spend
7. 7 Day Total Sales
8. 7 Day Total Orders (#)

### Business Data (Required for TACOS)
Export your Business Report from Amazon Seller Central with these columns:
1. Date
2. Ordered Product Sales
3. Units Ordered
4. Sessions

## Segment Classification

Campaigns are automatically classified into segments based on campaign name:
1. **Branded** - Campaign name contains "branded"
2. **Competitor** - Campaign name contains " pat " or "- pat -"
3. **Non-Branded** - All other campaigns

## Portfolio Classification

Portfolios are classified based on portfolio name:
1. **JN** - Portfolio name contains "JN"
2. **Non-JN** - All other portfolios

## Conditional Formatting

1. **Green highlights** - Positive changes (improvements)
2. **Red highlights** - Negative changes (declines)
3. For inverse metrics (ACoS, TACOS), lower is better so colors are reversed

## Files

| File | Description |
|------|-------------|
| `generate_excel_report.py` | Python script to generate the Excel template |
| `Campaign_Report_Template.xlsx` | Original Excel template (Desktop Excel) |
| `Campaign_Report_Template_v2.xlsx` | Excel Online-compatible version (recommended) |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

## Version Differences

| Feature | v1 (Original) | v2 (Excel Online) |
|---------|---------------|-------------------|
| Named Ranges | Yes | No (uses inline lists) |
| Pre-filled formulas in data sheets | Yes | No (shows formula text only) |
| Excel Online compatible | No | Yes |
| Desktop Excel compatible | Yes | Yes |

## Technical Details

1. Built with Python 3 and openpyxl
2. Uses Excel formulas (SUMIFS, IF, etc.) for calculations
3. Data validation dropdowns for interactive controls
4. Conditional formatting for visual indicators
5. Named ranges for maintainability

## Credits

Made by Krell for Piping Rock

Based on the HTML Campaign Report Generator from the automated-report project.
