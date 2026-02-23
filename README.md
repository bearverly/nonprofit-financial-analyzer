# Nonprofit Financial Statement Analyzer

A Streamlit application that analyzes bank statements and generates financial statements for nonprofit organizations.

## Features

- **Bank Statement Import**: Upload CSV or Excel bank statements
- **Smart Categorization**: Auto-categorize transactions into nonprofit-specific categories (Program Services, Management & General, Fundraising, Donations, Grants, etc.)
- **Financial Statements**: Generate FASB-compliant nonprofit financial statements:
  - Statement of Activities
  - Statement of Financial Position
  - Statement of Functional Expenses
  - Statement of Cash Flows
- **Dashboard**: Visual overview with charts and key metrics
- **Export**: Download financial statements as Excel workbooks

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Bank Statement Format

The app accepts CSV or Excel files. It will attempt to auto-detect columns, but works best with these columns:
- **Date**: Transaction date
- **Description**: Transaction description/memo
- **Amount**: Transaction amount (positive = deposit, negative = withdrawal) — or separate Debit/Credit columns

## Nonprofit Categories

### Revenue
- Donations & Contributions
- Grants
- Program Service Revenue
- Investment Income
- Fundraising Event Revenue
- Other Revenue

### Expenses
- Program Services
- Management & General
- Fundraising
- Facilities & Occupancy
- Salaries & Benefits
- Supplies & Materials
- Travel & Transportation
- Professional Services
- Insurance
- Other Expenses

## Usage

1. Upload your bank statement CSV/Excel file
2. Map columns if auto-detection doesn't match
3. Review and adjust auto-categorized transactions
4. Generate financial statements
5. Export to Excel
