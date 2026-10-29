# Veda Technology – Task 29: Basic Conditional Formatting

## Objective
Automatically highlight top and bottom values using Excel conditional formatting and improve report readability.

## Dataset
`raw_data/retail_sales.csv` contains 30 sample retail orders with Order ID, Order Date, Region, Category, Product Name, Quantity, Unit Price, Sales, and Profit.

## Tools
- Microsoft Excel-compatible `.xlsx` workbook
- Python / openpyxl used to create and verify the workbook

## Deliverables
1. `workbook/Retail_Sales_Conditional_Formatting.xlsx` – formatted workbook with **Retail Sales** and **Rules List** sheets.
2. `raw_data/retail_sales.csv` – source dataset.
3. `documentation/Rules_List.csv` – rule register.
4. `documentation/verify_workbook.py` – basic verification script.
5. `TASK_SUMMARY.txt` – task overview.

## Conditional formatting rules
- Sales: red-yellow-green color scale.
- Top 10% Sales: top-value conditional rule.
- Negative Profit: light red fill and dark red bold text.
- Profit >= 100: light green fill and dark green bold text.
- Quantity <= 3: light orange fill and bold text.
- Quantity: blue data bars.

## How to review
1. Open the workbook in Microsoft Excel.
2. On **Retail Sales**, review the Sales, Profit, and Quantity columns.
3. To inspect rules, select **Home → Conditional Formatting → Manage Rules**.
4. Open **Rules List** for target ranges, conditions, formatting, and purpose.
5. Test edge cases by changing Sales, Profit, or Quantity values; formatting updates automatically.

## Verification checklist
- [x] Workbook saved as `.xlsx`.
- [x] Source CSV and workbook contain the same 30 orders.
- [x] Conditional formatting rules are saved in the workbook.
- [x] Rules List worksheet is included.

**Note:** This is a practice dataset created for this task, not data from a real retailer.
