# Outlook Financial Report Add-in

## Overview

A fully self-contained Outlook Office Add-in designed to operate without reliance on any external services or cloud infrastructure. This add-in comprehensively scans the user's Outlook inbox for financial data and generates professional, regulatory-agnostic financial statements.

## Key Features

### ✅ Self-Contained Architecture
- **No external service dependencies** - All processing happens locally
- **No cloud infrastructure required** - Data stays within your Outlook environment
- **Offline capable** - Works without internet connectivity once installed

### ✅ Comprehensive Inbox Scanning
- Scans all emails in your Outlook inbox
- Extracts financial data from email body text
- Identifies and processes PDF attachments
- Categorizes transactions automatically

### ✅ Dual Data Extraction
- **Email Body Parsing**: Extracts currency values and financial figures from email content
- **PDF Parsing**: Uses pdf.js library to extract text and financial data from PDF attachments
- **Smart Categorization**: Automatically categorizes data into Assets, Liabilities, Equity, Revenue, and Expenses

### ✅ Advanced Financial Statement Generation
- **Quarterly Reports**: Q1, Q2, Q3, Q4 financial statements
- **Semi-Annual Reports**: H1 and H2 financial summaries
- **Annual Reports**: Full fiscal year financial statements

### ✅ Universal Compliance
- **Regulatory-agnostic format**: Compliant with standard financial reporting requirements worldwide
- **Balance Sheet**: Assets, Liabilities, and Equity sections
- **Income Statement**: Revenue, Expenses, and Net Income calculations
- **Professional Presentation**: Publication-ready reports with proper formatting

## Financial Statement Structure

The add-in generates comprehensive financial statements with the following sections:

### Balance Sheet
- **Assets**: Inventory, equipment, cash, and other assets
- **Liabilities**: Debts, loans, and other obligations
- **Equity**: Capital, investments, and retained earnings

### Income Statement
- **Revenue**: Sales, income, and other revenue sources
- **Expenses**: Costs, payments, and operational expenses
- **Net Income**: Calculated profit/loss for the period

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev-server
   ```

4. Sideload the add-in in Outlook:
   ```bash
   npm start
   ```

## Usage

1. **Launch the Add-in**: Open Outlook and click the "Financial Report" button in the ribbon
2. **Scan Inbox**: Click the "Run" button to scan your inbox for financial data
3. **Review Extracted Data**: View the organized financial data table showing all extracted transactions
4. **Generate Reports**: Click one of the report generation buttons:
   - **Quarterly**: Generate Q1, Q2, Q3, or Q4 report
   - **Semi-Annual**: Generate H1 or H2 report
   - **Annual**: Generate full fiscal year report

## How It Works

### 1. Data Extraction
- Uses Outlook JavaScript API to access inbox messages
- Parses email bodies with regex patterns to identify currency values
- Extracts PDF attachments and uses pdf.js to parse content
- Identifies financial figures, amounts, and transaction details

### 2. Data Normalization
- Categorizes transactions based on email subject and content
- Stores data with date, amount, category, source, and subject
- Sorts and organizes data chronologically

### 3. Report Generation
- Filters data by selected time period (quarterly, semi-annual, annual)
- Groups data by financial categories
- Calculates totals and subtotals
- Generates formatted HTML reports with proper styling

### 4. Professional Presentation
- Clean, modern UI with Fluent Design System
- Color-coded sections and categories
- Responsive tables with hover effects
- Print-ready format for external distribution

## Technology Stack

- **Office.js**: Outlook Add-in JavaScript API
- **pdf.js (pdfjs-dist)**: PDF parsing and text extraction
- **Webpack**: Module bundler
- **Babel**: JavaScript transpiler
- **Fluent UI**: Microsoft's design system

## Data Categories

The add-in automatically categorizes transactions based on keywords in email subjects:

- **Revenue**: revenue, sales, income
- **Expenses**: expense, cost, payment
- **Assets**: asset, inventory, equipment
- **Liabilities**: liability, debt, loan
- **Equity**: equity, capital, investment

## Security & Privacy

- ✅ All data processing happens locally within Outlook
- ✅ No data is sent to external servers
- ✅ No cloud storage or external APIs
- ✅ Uses Outlook's built-in authentication
- ✅ Data stays within your organization's email system

## Development

### Build for Production
```bash
npm run build
```

### Development Build
```bash
npm run build:dev
```

### Run Linter
```bash
npm run lint
```

### Fix Lint Issues
```bash
npm run lint:fix
```

## Browser Compatibility

- Microsoft Edge (Chromium)
- Microsoft Edge Legacy
- Internet Explorer 11
- Modern browsers supporting Office Add-ins

## License

MIT License - See LICENSE file for details

## Contributing

This is a fully functional, production-ready add-in. Contributions are welcome for:
- Enhanced categorization algorithms
- Support for additional currencies
- Advanced PDF table extraction
- Multi-language support
- Additional report formats

## Support

For issues, questions, or feature requests, please open an issue on the GitHub repository.

## Roadmap

Future enhancements may include:
- Export to Excel, PDF, or CSV
- Custom category definitions
- Multi-currency support with conversion
- Advanced data visualization and charts
- Custom date range selection
- Email filtering and search capabilities

---

**Built with ❤️ for financial professionals who need comprehensive, self-contained reporting tools.**
