# Personal Finance Dashboard

A local, privacy-focused HTML/JavaScript dashboard for visualizing personal finance data from an Excel file. No server required, no data leaves your computer.

## Features

### Net Worth & Forecast
- **Current Snapshot**: View total net worth, ISAs, liquid assets, and foreign holdings
- **Holdings Table**: See all active accounts with balances and interest rates
- **Historical Tracking**: Net worth growth over time with total and average monthly growth stats
- **Forecasting**: Project future net worth with compound interest calculations and optional monthly contributions

### Expense Analysis
- **Time Period Filters**: View data by specific month or use preset periods (Last 30 Days, Last 6 Months, Year to Date, All Time)
- **Category Filters**: Toggle categories on/off to focus your analysis. Transfers are excluded by default as they represent money movement rather than actual expenses.

### Summary Statistics
- **Total Income**: Sum of all income in the selected period
- **Total Expenses**: Sum of all expenses (excluding transfers)
- **Net Savings**: Income minus expenses
- **Savings Rate**: Percentage of income saved
- **Average Monthly Spending**: Average expenses per month

### Visualizations
1. **Monthly Income vs Expenses** (Line Chart): Track your financial trends over time
2. **Expenses by Category** (Pie Chart): See where your money goes
3. **Top Spending Categories** (Bar Chart): Identify your biggest expense categories

## Quick Start

1. **View the Dashboard with Dummy Data**:
   - Start a local server: `python3 -m http.server 8000`
   - Open `http://localhost:8000` in your browser
   - The dashboard will load with dummy data from `data_dummy/finance.2025.xlsx`

2. **Use Your Own Data**:
   - Create `data/finance.2025.xlsx` with your own data (see format below)
   - The dashboard will automatically detect and use your data instead of the dummy data

## Project Structure

```
finance-dashboard/
├── index.html              # Main dashboard page
├── style.css              # Styling and layout
├── app.js                 # Data processing and chart logic
├── data/
│   └── finance.2025.xlsx   # Your personal data (optional - auto-loads if exists)
├── data_dummy/
│   └── finance.2025.xlsx   # Dummy data (loads by default)
├── .gitignore             # Protects your personal data from git
└── README.md              # This file
```

## Data Format

### Excel File: `data/finance.2025.xlsx`

The Excel file must contain **three sheets** with the following names and structures:

#### Sheet 1: "Transactions"

Required columns:
- **Order**: Sequential number for each transaction
- **Date**: Transaction date in DD-MMM-YY format (e.g., "22-Jan-25")
- **Transaction Type**: "Purchase" or "Credit"
- **Transaction Description**: Description/merchant name
- **Amount**: Numeric value - negative for expenses, positive for income (e.g., -7.42, 3500)
- **Balance**: Running balance after transaction (numeric, no £ symbol)
- **Category**: Category name (must match categories from Categories sheet)

Example:
```
Order | Date      | Transaction Type | Transaction Description | Amount  | Balance | Category
1     | 22-Jan-25 | Purchase        | Waitrose                | -7.42   | 89.49   | Groceries
2     | 23-Jan-25 | Purchase        | TFL                     | -8.50   | 80.99   | Transport
```

#### Sheet 2: "Categories"

Required columns:
- **Keyword**: Text to match in transaction descriptions (case-insensitive)
- **Category**: Category to assign
- **Occurrence**: Count of matches (optional, can be left empty)
- **Comment**: Optional notes
- **Where**: Optional location/context

Example:
```
Keyword | Category   | Occurrence | Comment | Where
KFC     | Eating Out | 8          |         |
TESCO   | Groceries  | 24         |         |
```

#### Sheet 3: "Snapshots"

Required columns:
- **Date**: Snapshot date in DD-MMM-YY or DD/MM/YYYY format (e.g., "18-Sep-23" or "18/09/2023")
- **Account_Name**: Name of the account
- **Balance**: Numeric balance (no currency symbols)
- **Currency**: Currency code (e.g., GBP, AUD, USD)
- **Interest_Rate**: Annual interest rate as number (e.g., 4.5 for 4.5%), can be empty
- **Rate_Type**: Type of rate - AER, Gross, or PA (can be empty)
- **Notes**: Optional notes about the account

Example:
```
Date      | Account_Name | Balance | Currency | Interest_Rate | Rate_Type | Notes
18-Sep-23 | Chase        | 3000    | GBP      |               |           |
18-Sep-23 | Savings ISA  | 15000   | GBP      | 5.17          | AER       | Fixed until Dec 2025
```

## Important Notes

### Date Formats Supported

The dashboard automatically handles multiple date formats:

**Transactions Sheet**:
- Excel format: DD-MMM-YY (e.g., "22-Jan-25")
- This is the standard format when dates are entered in Excel

**Snapshots Sheet**:
- Excel format: DD-MMM-YY (e.g., "18-Sep-23")
- Alternative: DD/MM/YYYY (e.g., "18/09/2023")
- Excel serial dates are automatically converted

**Amounts**:
- Use plain numbers without currency symbols (e.g., -7.42, 3500)
- Negative for expenses, positive for income
- Excel will format these automatically if you apply currency formatting

### Currency Conversion

The dashboard supports multi-currency accounts:
- Foreign currency balances are automatically converted to GBP
- Exchange rates are fetched from exchangerate-api.com
- If API is unavailable, fallback rates are used (AUD: 0.52)

### Recommended Categories

- **Income**: Salary, bonuses, refunds
- **Groceries**: Supermarket shopping
- **Eating Out**: Restaurants, takeaways, coffee shops
- **Transport**: Public transport, taxis, fuel
- **Shopping**: Clothing, general retail
- **Entertainment**: Cinema, concerts, events
- **Health**: Pharmacy, medical expenses
- **Personal Care**: Grooming, beauty
- **Bills/Subscriptions**: Gym, streaming services, utilities, phone
- **Tech/Services**: Cloud services, domains, software
- **Housing**: Rent, mortgage
- **Transfers**: Money movements between accounts (excluded from expense calculations)

## How to Use Your Own Data

The dashboard comes with dummy data in `data_dummy/finance.2025.xlsx` so you can see how it works out of the box. When you're ready to use your own data:

### Step 1: Create Your Excel File

1. Create a new Excel file named `finance.2025.xlsx` in the `data/` folder (not data_dummy)
2. Create three sheets named exactly: **Transactions**, **Categories**, **Snapshots**
3. Add the column headers as specified in the Data Format section above
4. The dashboard will automatically detect and load your data instead of the dummy data

### Step 2: Add Your Transaction Data

1. Export transactions from your bank (usually as CSV)
2. Copy the data into the "Transactions" sheet
3. Ensure dates are in DD-MMM-YY format (Excel will format these automatically)
4. Ensure amounts are plain numbers (negative for expenses, positive for income)
5. Add a Category column (can be blank initially - will be auto-categorized)

### Step 3: Set Up Categories

1. In the "Categories" sheet, add keywords that appear in your transaction descriptions
2. Map each keyword to a category (e.g., "TESCO" → "Groceries")
3. The dashboard will automatically match keywords and categorize transactions
4. Leave Occurrence column blank - it will be calculated automatically

### Step 4: Add Net Worth Snapshots (Optional)

1. In the "Snapshots" sheet, add rows for each account snapshot
2. Take snapshots monthly or at regular intervals to track net worth over time
3. Include interest rates for accounts that earn interest
4. Specify Rate_Type as AER, Gross, or PA

### Tips for Better Categorization

- Use specific keywords first (e.g., "TESCO EXTRA" before "TESCO")
- Include common variations (e.g., "SAINSBURY", "SAINSBURYS")
- Keywords are matched case-insensitively
- Start with broad categories and refine over time
- Add new categories as needed - the dashboard will update automatically

## Privacy & Security

### Your Data Stays Local
- All processing happens in your browser
- No data is sent to any server (except for exchange rate API)
- No analytics or tracking
- Works completely offline (after first load of Chart.js and SheetJS from CDN)

### Git Protection
The `.gitignore` file is configured to prevent your personal financial data from being committed to git:
- All Excel files (`*.xlsx`, `*.xls`) in `data/` are ignored by default
- Only template/example files will be tracked (e.g., `*.example.xlsx`)

**Important**: If you plan to version control this project, never commit your personal Excel file containing financial data.

## Customization

### Adding More Charts
Edit `app.js` and use Chart.js documentation: https://www.chartjs.org/

### Changing Colors
Modify CSS variables in `style.css`:
```css
:root {
    --primary-blue: #2563eb;
    --success-green: #10b981;
    --danger-red: #ef4444;
    /* ... more colors ... */
}
```

### Adding New Categories
Simply add rows to the "Categories" sheet in your Excel file with new keywords and category names. Save the file and refresh the dashboard to see updates.

## Troubleshooting

### Dashboard Shows No Data
- Ensure `finance.2025.xlsx` exists in the `data/` folder
- Check browser console for errors (F12)
- Verify the Excel file has three sheets named exactly: "Transactions", "Categories", "Snapshots"
- Ensure column headers match the required format exactly

### Charts Not Displaying
- Check that transactions have valid dates (Excel will usually format these correctly)
- Ensure Amount column contains plain numbers (positive for income, negative for expenses)
- Remove any currency symbols (£, $) from the Amount and Balance columns
- Try refreshing the page

### Categories Not Working
- Ensure keywords in the Categories sheet match text in your transaction descriptions
- Keywords are case-insensitive
- Check for extra spaces or typos in category names
- The Category column in Transactions can be left blank - it will be auto-populated based on keywords

### Net Worth Section Not Showing
- Ensure the "Snapshots" sheet exists in your Excel file
- Check that dates are formatted correctly
- Verify Balance column contains numbers without currency symbols
- Interest_Rate should be a number (e.g., 5.17 for 5.17%)

### "Cannot load Excel file" Error
- You need to open the dashboard via a local web server or by opening the HTML file directly
- Some browsers block file:// protocol requests for security. If this happens, use a simple local server:
  ```bash
  # Python 3
  python -m http.server 8000

  # Then visit http://localhost:8000
  ```
- Ensure the Excel file is not open in Excel while trying to load it in the browser

## Technology Stack

- **Pure HTML/CSS/JavaScript**: No build process required
- **Chart.js**: Beautiful, responsive charts (loaded from CDN)
- **SheetJS (XLSX)**: Fast Excel file parsing (loaded from CDN)
- **No frameworks**: Lightweight and fast

## Future Enhancements

Ideas for extending the dashboard:
- Export filtered data to CSV/Excel
- Add budget tracking and warnings
- Compare multiple years
- Add tags/notes to transactions
- Import from multiple banks
- Recurring transaction detection
- Investment portfolio tracking
- Tax year analysis
- Automatic categorization using ML

## License

This is a personal project. Feel free to modify and use as needed.

## Support

This dashboard is designed to work out of the box. If you encounter issues:
1. Check the Troubleshooting section above
2. Verify your Excel file has the three required sheets with correct column headers
3. Check browser console for error messages (press F12)
4. Ensure you're opening the dashboard via a local server or directly in the browser

---

**Remember**: Keep your financial data private. Never share your actual Excel file or commit it to public repositories.
