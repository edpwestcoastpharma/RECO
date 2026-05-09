# Ledger Reconciliation Web

Local web software for account-team ledger reconciliation without any AI/LLM dependency.

## Run

```powershell
npm install
npm start
```

Open:

```text
http://localhost:3030
```

## Use

- Upload one company ledger PDF.
- Upload one or many party ledger PDFs.
- Upload the Excel template once; after that the app reuses `data/templates/ledger-reconciliation-template.xlsx`.
- The app preserves the template structure and writes reconciliation rows into the workbook.

## Supported Ledger Types

- Tally company ledgers with purchase/TDS and bank payment vouchers.
- Tally party ledgers.
- Party account statements with sale/bill lines.
- Odoo partner ledgers with multiple party PDFs such as G/V statements.

The backend creates a JSON audit report with extracted invoices, payments, ADD/LESS rows, TDS, and final verification.
