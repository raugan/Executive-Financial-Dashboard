# Financial Dashboard - Quick Usage Guide

## Getting Started in 5 Minutes

### Step 1: Open the Dashboard

Open **Financial_Dashboard.xlsx** and navigate to the **Dashboard** sheet. You will see a comprehensive financial overview with KPI cards, charts, and executive insights.

### Step 2: Select Your Quarter

Look at the top of the dashboard where you will find **"SELECT QUARTER:"** in cell B4. Click on the dropdown in cell C4 and choose the quarter you want to analyze: Q1 2025, Q2 2025, Q3 2025, or Q4 2025.

The moment you select a quarter, the entire dashboard updates automatically with that quarter's financial performance metrics.

### Step 3: Review KPI Cards

The dashboard displays eight key performance indicators in card format:

**Top Row:**
- Current Ratio (measures ability to pay short-term obligations)
- Quick Ratio (measures immediate liquidity without inventory)
- Gross Margin (percentage of revenue after cost of goods sold)
- Net Profit Margin (percentage of revenue that becomes profit)

**Bottom Row:**
- ROE (Return on Equity - profitability relative to shareholder equity)
- Debt-to-Equity (financial leverage and solvency)
- Interest Coverage (ability to service debt obligations)
- Asset Growth (quarter-over-quarter growth rate)

### Step 4: Analyze the Charts

The dashboard provides four visualization components:

**Revenue Trend** (Top Left): Shows revenue progression across all four quarters, helping identify growth patterns or seasonal variations.

**Net Income vs Operating Expenses** (Top Right): Compares profitability against operational costs, revealing efficiency trends.

**Asset vs Liability Composition** (Bottom Left): Displays the balance sheet structure, showing how assets compare to liabilities each quarter.

**Financial Health Radar** (Bottom Right): Presents a comprehensive snapshot of Q4 financial health across five dimensions, making it easy to spot strengths and areas for improvement.

### Step 5: Read Executive Insights

At the bottom of the dashboard, you will find a dynamically generated summary that provides context for the selected quarter's performance. This narrative automatically updates based on your quarter selection and includes:

- Total revenue achieved
- Net profit margin percentage
- Current debt-to-equity ratio
- Current ratio
- Interest coverage assessment (strong vs adequate)

This text is designed to be board-ready and can be copied directly into presentations or reports.

## Updating the Dashboard with New Data

### Adding New Quarterly Data

When you receive new financial statements for subsequent quarters, follow this process:

**Step 1: Update Income Statement**
1. Go to the **Table_Income** sheet
2. Add a new row below the last quarter (after row 7)
3. Enter the new quarter label (for example, Q1 2026) in column A
4. Fill in all financial metrics across columns B through L
5. Ensure all numbers are formatted as currency without manual formatting symbols

**Step 2: Update Balance Sheet**
1. Go to the **Table_Balance** sheet
2. Add a new row below the last quarter (after row 7)
3. Enter the same quarter label in column A
4. Fill in all balance sheet items across columns B through U
5. Verify that assets equal liabilities plus equity

**Step 3: Extend Formulas**
1. Return to the **Dashboard** sheet
2. Update all formula ranges that reference rows 4-7 to include the new row (for example, change $4:$7 to $4:$8)
3. Add the new quarter to the dropdown list in cell C4 by editing the data validation

**Step 4: Verify Calculations**
1. Select the new quarter from the dropdown
2. Verify that all KPI cards show values (not dashes or errors)
3. Check that charts update to include the new data point
4. Review the Executive Insights text for accuracy

## Exporting for Board Presentations

### Using the VBA Macro (Recommended)

If you have installed the VBA macro (see installation instructions in the main README):

1. Press Alt+F8 to open the Macro dialog
2. Select **ExportDashboardToPDF**
3. Click Run
4. Choose where to save the PDF file
5. The PDF will open automatically after export

The macro ensures proper formatting: landscape orientation, one page fit, centered layout, and professional margins.

### Manual PDF Export

If you prefer not to use macros:

1. Ensure the Dashboard sheet is active
2. Click File > Print (or press Ctrl+P)
3. Select **Microsoft Print to PDF** as the printer
4. In Page Setup, configure:
   - Orientation: Landscape
   - Scaling: Fit to 1 page wide by 1 page tall
   - Margins: Narrow (0.5 inches)
   - Options: Uncheck gridlines and row/column headings
5. Click Print and choose save location

## Common Questions

**Q: Why do some KPIs show a dash (-) instead of a number?**

A: This indicates that the selected quarter does not have complete data in the source sheets. Verify that both the Income Statement and Balance Sheet have entries for that quarter.

**Q: Can I change the company name in the header?**

A: Yes. Click on cell A1 in the Dashboard sheet and edit the text "AURELIUS STRATEGIC HOLDINGS" to your company name. The formatting will remain intact.

**Q: How do I hide or show the helper columns (H, I, Z)?**

A: Right-click on the column header and select "Unhide" to show hidden columns, or "Hide" to conceal them. These columns contain supporting formulas that power the dashboard but do not need to be visible for normal use.

**Q: Can I add more quarters beyond Q4 2025?**

A: Absolutely. Follow the "Adding New Quarterly Data" instructions above. The system is designed to accommodate unlimited quarters as long as you extend the formula ranges appropriately.

**Q: What if I want to compare two quarters side by side?**

A: Currently, the dashboard shows one quarter at a time. For comparison, use the charts which show all quarters simultaneously. You can also export PDFs for two different quarters and place them side by side in a presentation.

**Q: Why does Asset Growth show a dash for Q1?**

A: Asset Growth is calculated quarter-over-quarter, so Q1 does not have a previous quarter to compare against within the current dataset. This is expected behavior. If you add Q4 2024 data, Q1 2025 growth will calculate automatically.

## Troubleshooting

**Problem: All KPIs show dashes**

Solution: Check that cell C4 contains a valid quarter name that exactly matches the quarter labels in your data sheets (including spaces and capitalization).

**Problem: Charts are blank**

Solution: Verify that the chart data ranges (rows 16-22 in the Dashboard sheet) contain formulas, not just blank cells. If formulas are missing, they may have been accidentally deleted.

**Problem: Executive Insights shows "Select a quarter to view insights"**

Solution: This message appears when no quarter is selected or when the helper formulas in column Z are not calculating. Check cell C4 and verify the helper formulas are intact.

**Problem: Numbers appear as text instead of formatted currency**

Solution: Select the cells, right-click, choose Format Cells, and apply the Custom format: $#,##0;($#,##0);-

## Best Practices for Board Presentations

1. Always select Q4 (or the most recent quarter) before exporting to PDF for board meetings
2. Review the Executive Insights text and customize if needed for specific board context
3. Export one PDF with the dashboard overview, then prepare separate detailed sheets if questions arise
4. Keep a backup copy of the Excel file before making any structural changes
5. Use consistent quarter naming (Q1 2025, not 2025-Q1 or Q1-25) to ensure formulas work correctly

## Need Help?

For technical questions about formulas, data structure, or customization, refer to the comprehensive README.md file included in this repository.

For questions about financial metrics or interpretation of the KPIs, consult with your CFO or financial controller.

---

This dashboard is designed to save you hours of manual report preparation while providing instant insights into your company's financial performance. Use it regularly to track trends, identify areas for improvement, and communicate financial health to stakeholders.
