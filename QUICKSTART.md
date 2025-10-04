# Quick Start Guide

## Getting Started in 5 Minutes

### Prerequisites
- Node.js (v14 or higher)
- npm (comes with Node.js)
- Outlook Desktop, Outlook on the Web, or Outlook Mac
- Microsoft 365 account

### Step 1: Install Dependencies
```bash
cd outlook-financial-report-addin
npm install
```

### Step 2: Start Development Server
```bash
npm run dev-server
```

The server will start on `https://localhost:3000`

### Step 3: Sideload the Add-in

#### For Outlook Desktop (Windows)
```bash
npm start
```
This will automatically sideload the add-in in Outlook Desktop.

#### For Outlook on the Web
1. Go to [Outlook on the Web](https://outlook.office.com)
2. Click Settings (gear icon) → View all Outlook settings
3. Go to General → Manage Add-ins
4. Click "+ Add from file"
5. Browse to your project folder and select `manifest.json`
6. Click "Install"

### Step 4: Use the Add-in

1. **Open an Email**: Click on any email in your inbox
2. **Launch Add-in**: Look for the "Financial Reports" button in the ribbon
3. **Click "Generate Financial Report"**: This opens the task pane
4. **Click "Run"**: Scans your inbox for financial data
5. **View Results**: See extracted data in the organized table
6. **Generate Report**: Click "Quarterly", "Semi-Annual", or "Annual" button

## Sample Workflow

### Scenario: Generate Q4 Financial Report

1. Ensure you have emails with financial data in your inbox
   - Subject examples: "Invoice Payment $5,000", "Revenue Report - $10,000"
   - PDF attachments with financial figures

2. Open Outlook and click "Generate Financial Report" button

3. Click "Run" to scan inbox
   - Wait for scan to complete
   - View extracted data in table

4. Click "Quarterly" button
   - View Q4 financial statement
   - See Balance Sheet and Income Statement sections
   - Review Net Income calculation

5. Print or export report as needed

## Testing with Sample Data

### Create Test Emails

1. **Revenue Email**
   ```
   Subject: Sales Revenue - October
   Body: Total revenue for October: $15,000.50
   ```

2. **Expense Email**
   ```
   Subject: Office Supplies Expense
   Body: Paid $1,200.00 for office supplies
   ```

3. **Asset Email**
   ```
   Subject: New Equipment Purchase
   Body: Purchased new computer equipment: $3,500.00
   ```

4. **Liability Email**
   ```
   Subject: Business Loan Payment
   Body: Loan payment due: $2,000.00
   ```

5. **Equity Email**
   ```
   Subject: Capital Investment
   Body: New investor capital: $25,000.00
   ```

### Expected Results

After scanning, you should see:
- 5 transactions in the data table
- Transactions categorized automatically
- Totals calculated correctly

When generating a report:
- Balance Sheet with Assets, Liabilities, Equity
- Income Statement with Revenue and Expenses
- Net Income calculation (Revenue - Expenses)

## Common Configuration

### Increase Email Scan Limit

Edit `src/taskpane/taskpane.js`:

```javascript
// Change from:
const restUrl = mailbox.restUrl + "/v2.0/me/mailfolders/inbox/messages?$top=10&$select=Subject,Body,HasAttachments";

// To (for 50 emails):
const restUrl = mailbox.restUrl + "/v2.0/me/mailfolders/inbox/messages?$top=50&$select=Subject,Body,HasAttachments";
```

### Add Custom Currency Support

Edit the regex in `extractFinancialData()`:

```javascript
// For Euro support:
const currencyRegex = /[$€][0-9,]+(\.[0-9]{2})?/g;

// For multiple currencies:
const currencyRegex = /[$€£¥][0-9,]+(\.[0-9]{2})?/g;
```

### Customize Categories

Edit `categorizeTransaction()` in `taskpane.js`:

```javascript
function categorizeTransaction(amountStr, subject) {
    const subjectLower = subject.toLowerCase();
    
    // Add your custom keywords
    if (subjectLower.includes('subscription') || subjectLower.includes('recurring')) {
        return 'Recurring Expenses';
    }
    
    // ... rest of the function
}
```

## Development Commands

### Build for Production
```bash
npm run build
```
Creates optimized build in `dist/` folder

### Development Build with Watch
```bash
npm run watch
```
Automatically rebuilds on file changes

### Lint Code
```bash
npm run lint
```
Checks code for errors and style issues

### Fix Lint Issues
```bash
npm run lint:fix
```
Automatically fixes fixable lint issues

### Validate Manifest
```bash
npm run validate
```
Validates the manifest.json file

### Stop Debugging
```bash
npm run stop
```
Stops the Outlook debugging session

## Debugging Tips

### Enable Console Logging

Add logging to `taskpane.js`:

```javascript
console.log("Scanning inbox...");
console.log("Found data:", financialData);
console.log("Financial store:", financialDataStore);
```

### View Console Output

- **Outlook Desktop**: Right-click in task pane → Inspect
- **Outlook Web**: F12 Developer Tools
- **VS Code**: Use debugging configuration in `.vscode/launch.json`

### Common Issues

**Issue**: Add-in doesn't load
- **Solution**: Check if dev server is running on localhost:3000
- **Solution**: Try clearing Outlook cache and reloading

**Issue**: No data extracted
- **Solution**: Verify emails contain currency symbols ($)
- **Solution**: Check browser console for errors

**Issue**: PDF parsing fails
- **Solution**: Ensure PDF is text-based (not scanned image)
- **Solution**: Check PDF.js worker URL in console

## Production Deployment

### Step 1: Build for Production
```bash
npm run build
```

### Step 2: Update Manifest URL
Edit `manifest.json` and change URLs from `localhost:3000` to your production domain:

```json
"code": {
    "page": "https://your-domain.com/taskpane.html"
}
```

### Step 3: Deploy Files
Upload the contents of `dist/` folder to your web server:
- taskpane.html
- taskpane.js
- commands.html
- commands.js
- assets/*

### Step 4: Update webpack.config.js
Change the production URL:

```javascript
const urlProd = "https://your-domain.com/";
```

### Step 5: Deploy Manifest
- Upload `manifest.json` to your organization's add-in catalog
- Or distribute via AppSource

## Next Steps

1. **Customize for Your Needs**: Modify categories, currencies, and report formats
2. **Add More Features**: Implement export, charts, or custom date ranges
3. **Integrate with Systems**: Connect to accounting software or databases
4. **Train Users**: Provide training on categorization and report interpretation
5. **Monitor Usage**: Track adoption and gather feedback for improvements

## Support Resources

- **Documentation**: See `IMPLEMENTATION.md` for technical details
- **GitHub Issues**: Report bugs or request features
- **Office Add-ins Documentation**: [Microsoft Docs](https://docs.microsoft.com/office/dev/add-ins/)
- **PDF.js Documentation**: [Mozilla PDF.js](https://mozilla.github.io/pdf.js/)

## License

MIT License - Free for personal and commercial use

---

**Happy Financial Reporting! 📊💼**
