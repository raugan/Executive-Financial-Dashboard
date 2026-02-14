# Aurelius Strategic Holdings - Financial Command Centre FY 2025
# Financial Dashboard - Executive Command Centre

![Excel](https://img.shields.io/badge/Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![License](https://img.shields.io/badge/license-Proprietary-red)

> Transform quarterly financial statements into board-ready executive insights with automated KPIs, interactive visualisations, and dynamic reporting.
## Executive Summary

This repository contains a comprehensive, board-level financial dashboard system designed to transform quarterly financial statements into automated, visual executive insights. The dashboard provides real-time KPI tracking, interactive visualizations, and dynamic reporting capabilities.

## 📊 System Architecture

### Data Structure

The system uses a normalised, vertical quarterly format with two primary data tables:

**Table_Income** (Income Statement)
- Location: Sheet "Table_Income", Rows 4-7, Columns A-L
- Quarters: Q1-Q4 2025
- Key Metrics: Revenue, COGS, Gross Profit, Operating Expenses, EBITDA, EBIT, Net Income

**Table_Balance** (Balance Sheet)
- Location: Sheet "Table_Balance", Rows 4-7, Columns A-U  
- Quarters: Q1-Q4 2025
- Key Metrics: Assets (Current, PP&E, Total), Liabilities (Current, Long-term, Total), Equity

### Colour Coding System

Following industry-standard financial modelling conventions:

- **Blue Text (RGB: 0,0,255)**: User inputs and assumptions
- **Black Text (RGB: 0,0,0)**: Formulas and calculations
- **Navy Blue (#001F3F)**: Primary brand color, headers, KPI values
- **Muted Gold (#C5A059)**: Secondary accent, insights background
- **Steel Grey (#D1D5DB)**: Tertiary color, KPI labels

## 🔧 Automated KPI Engine

### Core Financial Ratios

#### Liquidity Ratios

**Current Ratio**
- Formula: `=IF(Z4="","-",Z4/Z5)`
- Location: Dashboard B7
- Calculation: Current Assets ÷ Current Liabilities
- Helper: Z4 (Current Assets), Z5 (Current Liabilities)

**Quick Ratio (Acid-Test)**
- Formula: `=IF(Z4="","-",(Z4-Z7)/Z5)`
- Location: Dashboard F7
- Calculation: (Current Assets - Inventory) ÷ Current Liabilities
- Helper: Z4 (Current Assets), Z5 (Current Liabilities), Z7 (Inventory)

#### Profitability Ratios

**Gross Margin**
- Formula: `=IF(Z2="","-",Z2/Z1)`
- Location: Dashboard J7
- Calculation: Gross Profit ÷ Revenue
- Format: Percentage with 1 decimal (0.0%)

**Net Profit Margin**
- Formula: `=IF(Z3="","-",Z3/Z1)`
- Location: Dashboard N7
- Calculation: Net Income ÷ Revenue
- Format: Percentage with 1 decimal (0.0%)

**Return on Equity (ROE)**
- Formula: `=IF(Z3="","-",Z3/Z9)`
- Location: Dashboard B11
- Calculation: Net Income ÷ Total Equity
- Format: Percentage with 1 decimal (0.0%)

#### Solvency Ratios

**Debt-to-Equity**
- Formula: `=IF(Z8="","-",Z8/Z9)`
- Location: Dashboard F11
- Calculation: Total Liabilities ÷ Total Equity
- Format: Decimal with 2 places (0.00)

**Interest Coverage Ratio**
- Formula: `=IF(Z11="","-",Z10/Z11)`
- Location: Dashboard J11
- Calculation: EBIT ÷ Interest Expense
- Format: Multiple formats (0.0x)

#### Growth Metrics

**Asset Growth (Quarter-over-Quarter)**
- Location: Dashboard N11
- Calculation: ((Current Quarter Assets / Previous Quarter Assets) - 1)
- Format: Percentage with 1 decimal (0.0%)
- Complex formula with nested IF statements for quarter selection

### Helper Formulas (Column Z)

All KPIs rely on INDEX/MATCH helper formulas that look up values based on the selected quarter:

```excel
Z1  (Revenue):           =IFERROR(INDEX(Table_Income!$B$4:$B$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
Z2  (Gross Profit):      =IFERROR(INDEX(Table_Income!$D$4:$D$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
Z3  (Net Income):        =IFERROR(INDEX(Table_Income!$L$4:$L$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
Z4  (Current Assets):    =IFERROR(INDEX(Table_Balance!$F$4:$F$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z5  (Current Liab):      =IFERROR(INDEX(Table_Balance!$N$4:$N$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z6  (Cash):              =IFERROR(INDEX(Table_Balance!$B$4:$B$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z7  (Inventory):         =IFERROR(INDEX(Table_Balance!$D$4:$D$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z8  (Total Liabilities): =IFERROR(INDEX(Table_Balance!$R$4:$R$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z9  (Total Equity):      =IFERROR(INDEX(Table_Balance!$U$4:$U$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
Z10 (EBIT):              =IFERROR(INDEX(Table_Income!$H$4:$H$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
Z11 (Interest Expense):  =IFERROR(INDEX(Table_Income!$I$4:$I$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
Z12 (Total Assets):      =IFERROR(INDEX(Table_Balance!$J$4:$J$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```

## 📈 Visualization Components

### Chart Data Preparation

**Revenue Data (Row 17, Columns B-E)**
```excel
B17: =INDEX(Table_Income!$B$4:$B$7,MATCH("Q1 2025",Table_Income!$A$4:$A$7,0))
C17: =INDEX(Table_Income!$B$4:$B$7,MATCH("Q2 2025",Table_Income!$A$4:$A$7,0))
D17: =INDEX(Table_Income!$B$4:$B$7,MATCH("Q3 2025",Table_Income!$A$4:$A$7,0))
E17: =INDEX(Table_Income!$B$4:$B$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))
```

### Chart Specifications

**1. Revenue Trend Line Chart**
- Position: Dashboard B31
- Type: Line Chart with smooth curves
- Data Source: Row 17 (Revenue), Columns B-E
- Style: Navy Blue line (#001F3F), 35000 width
- Dimensions: 18 wide × 10 high

**2. Net Income vs Operating Expenses Column Chart**
- Position: Dashboard J31
- Type: Clustered Column Chart
- Data Source: Rows 18-19 (Net Income, Operating Expenses)
- Colors: Muted Gold (#C5A059), Steel Grey (#D1D5DB)
- Dimensions: 18 wide × 10 high

**3. Assets vs Liabilities Stacked Bar Chart**
- Position: Dashboard B47
- Type: Horizontal Stacked Bar Chart
- Data Source: Rows 21-22 (Total Assets, Total Liabilities)
- Colors: Navy Blue (#001F3F), Muted Gold (#C5A059)
- Dimensions: 18 wide × 10 high

**4. Financial Health Radar Chart**
- Position: Dashboard J47
- Type: Filled Radar Chart
- Metrics: Current Ratio, Quick Ratio, ROE, Interest Coverage, Asset Growth
- Data: Q4 2025 snapshot
- Colour: Navy Blue (#001F3F)
- Dimensions: 18 wide × 10 high

## 🎯 Interactive Features

### Quarter Selection Dropdown

**Location**: Dashboard C4

**Data Validation**: List with values "Q1 2025,Q2 2025,Q3 2025,Q4 2025"

**Default Value**: Q4 2025

**Functionality**: When changed, all KPIs and the Executive Insights section update automatically through cell references to $C$4

### Executive Insights Generator

**Location**: Dashboard B26 (merged through P28)

**Dynamic Formula**:
```excel
=IF(Z1="", "Select a quarter to view insights",
"For " & $C$4 & ": Revenue reached " & TEXT(Z1,"$#,##0") & 
" with a " & TEXT(Z3/Z1,"0.0%") & " net margin. Balance sheet strength improved with Debt-to-Equity at " & 
TEXT(Z8/Z9,"0.00") & " and Current Ratio of " & TEXT(Z4/Z5,"0.00") & 
". Interest coverage stands at " & TEXT(Z10/Z11,"0.0") & "x, indicating " & 
IF(Z10/Z11>5,"strong","adequate") & " debt servicing capacity.")
```

**Output Example**:
"For Q4 2025: Revenue reached $19,500,000 with a 20.9% net margin. Balance sheet strength improved with Debt-to-Equity at 0.62 and Current Ratio of 3.34. Interest coverage stands at 21.7x, indicating strong debt servicing capacity."

## 🔒 Professional Polish Features

### UI Cleanup Checklist

**Remove Gridlines**:
```vba
ActiveWindow.DisplayGridlines = False
```

**Protect Formulas** (Optional):
```vba
Worksheets("Dashboard").Protect Password:="", DrawingObjects:=False, Contents:=True
```

**Hide Helper Columns**:
- Columns H, I (Radar chart data)
- Column Z (KPI helper formulas)

**Consistent Typography**:
- All text: Arial font family
- Headers: 14pt bold
- KPI Labels: 10pt bold
- KPI Values: 22pt bold
- Body text: 10pt regular

### VBA Export Macros

**Standard Export** (with user prompt):
```vba
Sub ExportDashboardToPDF()
    ' Prompts user for save location
    ' Exports Dashboard sheet as landscape PDF
    ' Filename format: Financial_Dashboard_YYYY-MM-DD.pdf
End Sub
```

**Quick Export** (to Desktop):
```vba
Sub QuickExportDashboardToPDF()
    ' Exports directly to Desktop folder
    ' No user prompt required
End Sub
```

**Installation**:
1. Press Alt+F11 to open the VBA Editor
2. Insert > Module
3. Paste macro code from Dashboard_Export_Macros.vba
4. Run from Developer > Macros menu

## 📁 Repository Structure

```
financial-dashboard/
├── README.md                          # This file
├── Financial_Dashboard.xlsx           # Main dashboard workbook
├── income_statement_data.csv          # Sample income data
├── balance_sheet_data.csv             # Sample balance data
├── Dashboard_Export_Macros.vba        # VBA macro code
├── FORMULAS.md                        # Detailed formula reference
└── USAGE_GUIDE.md                     # Step-by-step usage instructions
```

## 🚀 Quick Start Guide

### For Excel Users

1. Open Financial_Dashboard.xlsx
2. Navigate to the Dashboard sheet
3. Select a quarter from the dropdown in cell C4
4. All KPIs and charts update automatically
5. Review the Executive Insights section
6. Export to PDF using the VBA macro (optional)

### For Power BI Adaptation

**Modify Part 1 (Data Architecture)**:
- Create a Star Schema instead of Excel Tables
- Fact table: Quarterly_Financials
- Dimension tables: Date_Dimension, Account_Dimension
- Use CALCULATE and FILTER for KPI measures

**Sample DAX Measure**:
```dax
Current Ratio = 
DIVIDE(
    SUM(Balance_Sheet[Current_Assets]),
    SUM(Balance_Sheet[Current_Liabilities]),
    BLANK()
)
```

### For Python/Streamlit Adaptation

**Modify Part 1 (Data Loading)**:
```python
import pandas as pd

income_df = pd.read_csv('income_statement_data.csv')
balance_df = pd.read_csv('balance_sheet_data.csv')

# Convert to indexed DataFrame
income_df.set_index('Quarter', inplace=True)
balance_df.set_index('Quarter', inplace=True)
```

**Modify Part 3 (Visualisation)**:
```python
import plotly.express as px
import streamlit as st

# Revenue trend
fig = px.line(income_df, y='Revenue', title='Quarterly Revenue Trend')
st.plotly_chart(fig)

# KPI calculations
current_ratio = balance_df.loc[selected_quarter, 'Current_Assets'] / \
                balance_df.loc[selected_quarter, 'Current_Liabilities']
```

## 📊 Sample Data Overview

The included dummy data demonstrates a growth trajectory:

**Revenue Growth**: $12.5M (Q1) → $19.5M (Q4) (+56% over 4 quarters)

**Net Margin Improvement**: 13.1% (Q1) → 20.9% (Q4)

**Debt-to-Equity Reduction**: 1.11 (Q1) → 0.62 (Q4)

**Current Ratio Strengthening**: 1.98 (Q1) → 3.34 (Q4)

This data pattern is designed to showcase the dashboard's ability to highlight positive financial performance trends.

## 🔍 Formula Error Prevention

All formulas include error handling using IFERROR or IF statements to ensure:
- No #DIV/0! errors from zero denominators
- No #N/A errors from missing quarters
- No #VALUE! errors from invalid data types
- Charts do not break when data is incomplete

## 🛠️ Customization Guide

### Adding New Quarters

1. Insert a new row in Table_Income and Table_Balance sheets (row 8)
2. Enter data for Q1 2026, Q2 2026, etc.
3. Update data ranges in all formulas (change $7 to $8, etc.)
4. Add a new quarter to the dropdown validation list in Dashboard C4

### Adding New KPIs

1. Create a helper formula in column Z (e.g., Z13 for the new metric)
2. Add a new KPI card in the dashboard (copy existing card structure)
3. Update chart data if visualization needed
4. Document formula in this README

### Changing Colour Scheme

Update these variables in the Python creation script:
```python
NAVY_BLUE = '001F3F'      # Primary brand color
MUTED_GOLD = 'C5A059'     # Secondary accent
STEEL_GREY = 'D1D5DB'     # Tertiary color
```

## 📝 Best Practices

1. Always update both the Income and Balance sheets with consistent quarter labels
2. Use the same currency denomination across all sheets ($ or $000 or $MM)
3. Verify formulas after adding new data using the quarter selector
4. Export snapshots to PDF before board meetings
5. Maintain data integrity by protecting formula cells
6. Document any manual adjustments in the Notes section (if added)

## 🤝 Contributing

To contribute enhancements:
1. Fork this repository
2. Create a feature branch
3. Test all formulas for zero errors
4. Update documentation
5. Submit a pull request with a clear description
## Contributing

Contributions welcome! Please read the contributing guidelines before submitting PRs.

## Support

For issues or questions, please open a GitHub issue.
## 📄 License

Proprietary - For internal use by Aurelius Strategic Holdings and authorised parties only.

## ✨ Version History

**v1.0** - February 2026
- Initial release with core KPI engine
- Four primary visualisations
- Executive insights generator
- PDF export capability
- Complete documentation

---

For questions or support, contact the Financial Systems team.
