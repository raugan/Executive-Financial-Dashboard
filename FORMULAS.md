# Formula Reference Guide

## Complete Formula Documentation for Financial Dashboard

This document provides a comprehensive reference for every formula used in the Financial Command Centre, organized by component and purpose. Use this guide when customizing, troubleshooting, or extending the dashboard functionality.

## Helper Formula System (Column Z)

The dashboard uses a centralized helper formula system in column Z to retrieve data based on the selected quarter. All KPI calculations reference these helper cells, creating a single source of truth that updates when the user changes the quarter selection in cell C4.

### Quarter Lookup Formulas

Each helper formula follows this pattern: retrieve a specific financial metric for the quarter selected in C4 by matching against the quarter column in the source table and returning the corresponding value.

**Z1 - Revenue**
```
=IFERROR(INDEX(Table_Income!$B$4:$B$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
```
Returns the revenue value from column B of the Income Statement for the selected quarter. The IFERROR wrapper returns blank if the quarter is not found, preventing errors.

**Z2 - Gross Profit**
```
=IFERROR(INDEX(Table_Income!$D$4:$D$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
```
Returns gross profit from column D. This is used to calculate the gross margin percentage.

**Z3 - Net Income**
```
=IFERROR(INDEX(Table_Income!$L$4:$L$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
```
Returns net income from column L, the bottom line of the income statement. Used in multiple ratio calculations including net margin and ROE.

**Z4 - Current Assets**
```
=IFERROR(INDEX(Table_Balance!$F$4:$F$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns current assets from column F of the Balance Sheet. Critical component of liquidity ratios.

**Z5 - Current Liabilities**
```
=IFERROR(INDEX(Table_Balance!$N$4:$N$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns current liabilities from column N. Used as the denominator in liquidity ratios.

**Z6 - Cash and Equivalents**
```
=IFERROR(INDEX(Table_Balance!$B$4:$B$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns cash position from column B. Available for enhanced liquidity analysis if needed.

**Z7 - Inventory**
```
=IFERROR(INDEX(Table_Balance!$D$4:$D$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns inventory value from column D. Subtracted from current assets to calculate the quick ratio.

**Z8 - Total Liabilities**
```
=IFERROR(INDEX(Table_Balance!$R$4:$R$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns total liabilities from column R. Numerator for the debt-to-equity ratio.

**Z9 - Total Equity**
```
=IFERROR(INDEX(Table_Balance!$U$4:$U$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns total equity from column U. Denominator for debt-to-equity and ROE calculations.

**Z10 - EBIT (Earnings Before Interest and Taxes)**
```
=IFERROR(INDEX(Table_Income!$H$4:$H$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
```
Returns EBIT from column H. Numerator for interest coverage ratio.

**Z11 - Interest Expense**
```
=IFERROR(INDEX(Table_Income!$I$4:$I$7,MATCH($C$4,Table_Income!$A$4:$A$7,0)),"")
```
Returns interest expense from column I. Denominator for interest coverage ratio.

**Z12 - Total Assets**
```
=IFERROR(INDEX(Table_Balance!$J$4:$J$7,MATCH($C$4,Table_Balance!$A$4:$A$7,0)),"")
```
Returns total assets from column J. Used in asset growth calculation and available for additional analysis.

## KPI Card Formulas

Each KPI uses the helper formulas in column Z and includes error handling to display a dash when data is unavailable.

### Liquidity KPIs

**Current Ratio (Cell B7)**
```
=IF(Z4="","-",Z4/Z5)
```
Divides current assets by current liabilities. A ratio above two generally indicates strong liquidity. The IF statement displays a dash if the current assets helper cell is blank, which occurs when no quarter is selected or data is missing.

**Quick Ratio (Cell F7)**
```
=IF(Z4="","-",(Z4-Z7)/Z5)
```
Calculates the acid-test ratio by subtracting inventory from current assets before dividing by current liabilities. This provides a more conservative view of liquidity since inventory may not be quickly convertible to cash. A ratio above one typically indicates adequate short-term liquidity.

### Profitability KPIs

**Gross Margin (Cell J7)**
```
=IF(Z2="","-",Z2/Z1)
```
Expresses gross profit as a percentage of revenue. This shows how much of each revenue dollar remains after covering the direct costs of goods sold. Higher percentages indicate better pricing power or operational efficiency.

**Net Profit Margin (Cell N7)**
```
=IF(Z3="","-",Z3/Z1)
```
Expresses net income as a percentage of revenue, representing the bottom-line profitability. This metric accounts for all expenses including operating costs, interest, and taxes. Rising net margins indicate improving overall profitability.

**Return on Equity (Cell B11)**
```
=IF(Z3="","-",Z3/Z9)
```
Measures the return generated on shareholder equity. This shows how effectively the company uses shareholder capital to generate profits. Higher ROE percentages indicate more efficient use of equity capital.

### Solvency KPIs

**Debt-to-Equity (Cell F11)**
```
=IF(Z8="","-",Z8/Z9)
```
Compares total liabilities to total equity, indicating the degree of financial leverage. A ratio above one means the company has more debt than equity. Declining ratios indicate improving financial stability and reduced leverage.

**Interest Coverage Ratio (Cell J11)**
```
=IF(Z11="","-",Z10/Z11)
```
Measures how many times the company can cover its interest payments with operating earnings. A ratio above five typically indicates strong debt servicing capability, while ratios below two may signal potential difficulties in meeting interest obligations.

### Growth KPIs

**Asset Growth Quarter-over-Quarter (Cell N11)**
```
=IF($C$4="Q1 2025","",
IF($C$4="Q2 2025",
(INDEX(Table_Balance!$J$4:$J$7,MATCH("Q2 2025",Table_Balance!$A$4:$A$7,0))/
INDEX(Table_Balance!$J$4:$J$7,MATCH("Q1 2025",Table_Balance!$A$4:$A$7,0))-1),
IF($C$4="Q3 2025",
(INDEX(Table_Balance!$J$4:$J$7,MATCH("Q3 2025",Table_Balance!$A$4:$A$7,0))/
INDEX(Table_Balance!$J$4:$J$7,MATCH("Q2 2025",Table_Balance!$A$4:$A$7,0))-1),
IF($C$4="Q4 2025",
(INDEX(Table_Balance!$J$4:$J$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))/
INDEX(Table_Balance!$J$4:$J$7,MATCH("Q3 2025",Table_Balance!$A$4:$A$7,0))-1),
""))))
```

This complex nested IF formula calculates the percentage change in total assets from the previous quarter. For Q1, it returns blank since there is no prior quarter in the dataset. For each subsequent quarter, it retrieves the current quarter's assets and the previous quarter's assets, divides them, and subtracts one to express the result as a percentage change.

## Chart Data Formulas

The dashboard includes four charts, each pulling data from specific cells that contain formulas to extract quarterly data from the source sheets.

### Revenue Trend Data (Row 17, Columns B-E)

These formulas populate the data series for the line chart showing revenue trends across all quarters.

**B17 - Q1 Revenue**
```
=INDEX(Table_Income!$B$4:$B$7,MATCH("Q1 2025",Table_Income!$A$4:$A$7,0))
```

**C17 - Q2 Revenue**
```
=INDEX(Table_Income!$B$4:$B$7,MATCH("Q2 2025",Table_Income!$A$4:$A$7,0))
```

**D17 - Q3 Revenue**
```
=INDEX(Table_Income!$B$4:$B$7,MATCH("Q3 2025",Table_Income!$A$4:$A$7,0))
```

**E17 - Q4 Revenue**
```
=INDEX(Table_Income!$B$4:$B$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))
```

### Net Income and Operating Expenses Data (Rows 18-19, Columns B-E)

These formulas populate the clustered column chart comparing profitability against operational costs.

**Row 18 - Net Income for Q1-Q4**
Same pattern as revenue, but referencing column L (Net Income):
```
B18: =INDEX(Table_Income!$L$4:$L$7,MATCH("Q1 2025",Table_Income!$A$4:$A$7,0))
C18: =INDEX(Table_Income!$L$4:$L$7,MATCH("Q2 2025",Table_Income!$A$4:$A$7,0))
D18: =INDEX(Table_Income!$L$4:$L$7,MATCH("Q3 2025",Table_Income!$A$4:$A$7,0))
E18: =INDEX(Table_Income!$L$4:$L$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))
```

**Row 19 - Operating Expenses for Q1-Q4**
References column E (Operating Expenses):
```
B19: =INDEX(Table_Income!$E$4:$E$7,MATCH("Q1 2025",Table_Income!$A$4:$A$7,0))
C19: =INDEX(Table_Income!$E$4:$E$7,MATCH("Q2 2025",Table_Income!$A$4:$A$7,0))
D19: =INDEX(Table_Income!$E$4:$E$7,MATCH("Q3 2025",Table_Income!$A$4:$A$7,0))
E19: =INDEX(Table_Income!$E$4:$E$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))
```

### Balance Sheet Composition Data (Rows 21-22, Columns B-E)

These formulas populate the stacked bar chart showing assets versus liabilities.

**Row 21 - Total Assets for Q1-Q4**
```
B21: =INDEX(Table_Balance!$J$4:$J$7,MATCH("Q1 2025",Table_Balance!$A$4:$A$7,0))
C21: =INDEX(Table_Balance!$J$4:$J$7,MATCH("Q2 2025",Table_Balance!$A$4:$A$7,0))
D21: =INDEX(Table_Balance!$J$4:$J$7,MATCH("Q3 2025",Table_Balance!$A$4:$A$7,0))
E21: =INDEX(Table_Balance!$J$4:$J$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))
```

**Row 22 - Total Liabilities for Q1-Q4**
```
B22: =INDEX(Table_Balance!$R$4:$R$7,MATCH("Q1 2025",Table_Balance!$A$4:$A$7,0))
C22: =INDEX(Table_Balance!$R$4:$R$7,MATCH("Q2 2025",Table_Balance!$A$4:$A$7,0))
D22: =INDEX(Table_Balance!$R$4:$R$7,MATCH("Q3 2025",Table_Balance!$A$4:$A$7,0))
E22: =INDEX(Table_Balance!$R$4:$R$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))
```

### Financial Health Radar Data (Hidden Columns H-I)

The radar chart displays five metrics for Q4 2025 only, providing a snapshot of overall financial health.

**I17 - Current Ratio for Q4**
```
=INDEX(Table_Balance!$F$4:$F$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))/INDEX(Table_Balance!$N$4:$N$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))
```

**I18 - Quick Ratio for Q4**
```
=(INDEX(Table_Balance!$F$4:$F$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))-INDEX(Table_Balance!$D$4:$D$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0)))/INDEX(Table_Balance!$N$4:$N$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))
```

**I19 - ROE Percentage for Q4**
```
=INDEX(Table_Income!$L$4:$L$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))/INDEX(Table_Balance!$U$4:$U$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))*100
```

**I20 - Interest Coverage for Q4**
```
=INDEX(Table_Income!$H$4:$H$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))/INDEX(Table_Income!$I$4:$I$7,MATCH("Q4 2025",Table_Income!$A$4:$A$7,0))
```

**I21 - Asset Growth Percentage for Q4**
```
=(INDEX(Table_Balance!$J$4:$J$7,MATCH("Q4 2025",Table_Balance!$A$4:$A$7,0))/INDEX(Table_Balance!$J$4:$J$7,MATCH("Q3 2025",Table_Balance!$A$4:$A$7,0))-1)*100
```

## Executive Insights Formula

The executive insights section dynamically generates a narrative summary based on the selected quarter.

**Cell B26 (merged through P28)**
```
=IF(Z1="","Select a quarter to view insights",
"For " & $C$4 & ": Revenue reached " & TEXT(Z1,"$#,##0") & 
" with a " & TEXT(Z3/Z1,"0.0%") & " net margin. Balance sheet strength improved with Debt-to-Equity at " & 
TEXT(Z8/Z9,"0.00") & " and Current Ratio of " & TEXT(Z4/Z5,"0.00") & 
". Interest coverage stands at " & TEXT(Z10/Z11,"0.0") & "x, indicating " & 
IF(Z10/Z11>5,"strong","adequate") & " debt servicing capacity.")
```

This formula concatenates text with formatted values from the helper cells. It includes a nested IF statement that evaluates interest coverage: if the ratio exceeds five, it uses the word "strong," otherwise it uses "adequate." The TEXT function formats numbers with appropriate currency symbols, decimals, and thousands separators.

## Data Validation Formula

**Cell C4 - Quarter Selection Dropdown**

The quarter selector uses Excel's Data Validation feature with a list source:

```
"Q1 2025,Q2 2025,Q3 2025,Q4 2025"
```

This creates a dropdown allowing users to select from the four available quarters. When adding new quarters, this list must be updated to include the new options.

## Formula Maintenance Guidelines

When extending the dashboard with additional quarters, follow these update procedures:

### Adding Q1 2026

1. Update all helper formulas in column Z to expand the range from $4:$7 to $4:$8
2. Update all chart data formulas similarly
3. Add Q1 2026 to the dropdown validation list
4. Create new formulas for any additional chart data points

### Modifying Column Positions

If you rearrange columns in the source sheets, update the column references in all formulas. For example, if Revenue moves from column B to column C in Table_Income, update all instances of Table_Income!$B$ to Table_Income!$C$ throughout the dashboard.

### Error Handling Best Practices

All formulas use either IFERROR or IF statements to prevent calculation errors from disrupting the dashboard. When creating new formulas, always include error handling:

- Use IFERROR for lookup formulas that might not find a match
- Use IF to check for blank cells before performing divisions
- Return blank text ("") or dash ("-") rather than allowing errors to display

### Formula Testing Checklist

After modifying formulas, verify the following:

1. Select each quarter from the dropdown and confirm all KPIs calculate
2. Check that no cells display #DIV/0!, #VALUE!, #REF!, or #N/A errors
3. Verify charts update correctly with the new data
4. Confirm Executive Insights text generates properly
5. Test with edge cases (zero values, missing quarters)

## Additional Resources

For questions about specific formula syntax or Excel functions used in this dashboard:

- INDEX: Returns a value from a table based on row and column position
- MATCH: Finds the position of a value within a range
- IFERROR: Returns a specified value if a formula results in an error
- IF: Performs logical tests and returns different values based on the result
- TEXT: Formats numbers as text with specific formatting codes

Microsoft Excel documentation provides comprehensive guides for each of these functions.

---

This formula reference is maintained alongside the dashboard and should be updated whenever formulas are modified or extended.
