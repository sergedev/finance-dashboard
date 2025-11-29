# Personal Finance Dashboard

A local, privacy-focused HTML/JavaScript dashboard for visualizing personal finance data from CSV files. No server required, no data leaves your computer.

## Features

### Interactive Filtering
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

1. **View the Dashboard**: Simply open `index.html` in your web browser
2. **Explore Dummy Data**: The dashboard comes with a full year of realistic dummy data (2025)
3. **Replace with Your Data**: When ready, replace the CSV files in the `data/` folder with your own

## Project Structure

```
finance-dashboard/
├── index.html              # Main dashboard page
├── style.css              # Styling and layout
├── app.js                 # Data processing and chart logic
├── data/
│   ├── transactions.2025.csv   # Transaction data
│   └── categories.2025.csv     # Category keyword mappings
├── .gitignore             # Protects your personal data from git
└── README.md              # This file
```

## Data Format

### transactions.2025.csv

Required columns:
- **Order**: Sequential number for each transaction
- **Date**: Transaction date in DD-MMM-YY format (e.g., "07-Jan-25")
- **Transaction Type**: "Credit" or "Debit"
- **Transaction Description**: Description/merchant name
- **Amount**: Positive for income, negative for expenses (e.g., -17.14, 3500.00)
- **Balance**: Running balance after transaction
- **Category**: Category name (must match categories from categories.csv)

Example:
```csv
Order,Date,Transaction Type,Transaction Description,Amount,Balance,Category
1,01-Jan-25,Credit,Salary Payment,3500.00,3500.00,Income
2,02-Jan-25,Debit,TESCO EXTRA,-67.42,3432.58,Groceries
3,03-Jan-25,Debit,TFL TRAVEL CHARGE,-8.50,3424.08,Transport
```

### categories.2025.csv

Required columns:
- **Keyword**: Text to match in transaction descriptions (case-insensitive)
- **Category**: Category to assign
- **Occurrence**: Count of matches (can be calculated or left empty)
- **Comment**: Optional notes
- **Where**: Optional location/context

Example:
```csv
Keyword,Category,Occurrence,Comment,Where
TESCO,Groceries,24,,
SAINSBURY,Groceries,24,,
TFL,Transport,72,,
NETFLIX,Bills/Subscriptions,12,,
```

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

### Option 1: Export from Your Bank

1. Export transactions from your bank as CSV
2. Reformat to match the required column structure
3. Add Category column (can be blank initially)
4. Replace `data/transactions.2025.csv` with your file

### Option 2: Manual Categorization

1. Create a `categories.csv` file with keywords from your transactions
2. Map each keyword to a category
3. The dashboard will automatically categorize transactions based on keyword matches

### Tips for Better Categorization

- Use specific keywords first (e.g., "TESCO EXTRA" before "TESCO")
- Include common variations (e.g., "SAINSBURY", "SAINSBURYS")
- Review the Occurrence column to ensure keywords are matching correctly
- Start with broad categories and refine over time

## Privacy & Security

### Your Data Stays Local
- All processing happens in your browser
- No data is sent to any server
- No analytics or tracking
- Works completely offline (after first load of Chart.js and PapaParse from CDN)

### Git Protection
The `.gitignore` file is configured to prevent your personal financial data from being committed to git:
- All CSV files in `data/` are ignored by default
- Only template/example files will be tracked

**Important**: If you plan to version control this project, never commit your personal CSV files.

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
Simply add rows to `categories.csv` with new keywords and category names. Refresh the dashboard to see updates.

## Troubleshooting

### Dashboard Shows No Data
- Ensure CSV files are in the `data/` folder
- Check browser console for errors (F12)
- Verify CSV files are properly formatted with correct headers

### Charts Not Displaying
- Check that transactions have valid dates in DD-MMM-YY format
- Ensure Amount column contains valid numbers (positive for income, negative for expenses)
- Try refreshing the page

### Categories Not Working
- Verify category names in transactions match those in categories.csv
- Check for extra spaces or typos in category names
- Keywords are case-insensitive

### "Cannot load CSV files" Error
- You need to open the dashboard via a local web server or by opening the HTML file directly
- Some browsers block file:// protocol requests for security. If this happens, use a simple local server:
  ```bash
  # Python 3
  python -m http.server 8000

  # Then visit http://localhost:8000
  ```

## Technology Stack

- **Pure HTML/CSS/JavaScript**: No build process required
- **Chart.js**: Beautiful, responsive charts (loaded from CDN)
- **PapaParse**: Fast CSV parsing (loaded from CDN)
- **No frameworks**: Lightweight and fast

## Future Enhancements

Ideas for extending the dashboard:
- Export filtered data to CSV
- Add budget tracking and warnings
- Compare multiple years
- Forecast future spending
- Add tags/notes to transactions
- Multi-currency support
- Import from multiple banks
- Recurring transaction detection

## License

This is a personal project. Feel free to modify and use as needed.

## Support

This dashboard is designed to work out of the box. If you encounter issues:
1. Check the Troubleshooting section above
2. Verify your CSV files match the required format
3. Check browser console for error messages

---

**Remember**: Keep your financial data private. Never share your actual CSV files or commit them to public repositories.
