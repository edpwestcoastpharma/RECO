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
- Review diagnostics, ADD/LESS/PAYMENT/ADJUSTMENT rows, repair suggestions, and the generated Excel sheet before downloading Excel.

## Supported Ledger Types

- Tally company ledgers with purchase/TDS and bank payment vouchers.
- Tally party ledgers.
- Party account statements with sale/bill lines.
- Odoo partner ledgers with multiple party PDFs such as G/V statements.
- Old numeric Tally ledgers, including financial-year invoice keys.
- West Coast company-specific write-off rows when the source ledger contains a real `SUNDRY BALANCE WRITTEN OFF - NET` line.

The backend creates a JSON audit report with extracted invoices, payments, ADD/LESS rows, TDS, and final verification.

## Diagnostics And Review

Every job includes:

- parser diagnostics: detected ledger types, periods, counts, duplicate keys, warnings
- review rows: ADD, LESS, PAYMENT, and company adjustment previews
- rule-based repair: required ADD total, current ADD total, H63 gap, and likely next checks
- final Excel preview: browser grid rendered from the generated `.xlsx` output
- company rule library: supported West Coast ledger patterns and safety rules

`H63 = 0` means the itemized rows and allowed source-ledger adjustments explain the balance. If the app shows `NEEDS_REVIEW`, do not use the Excel as final until the mismatch is traced to real ledger rows.

## Regression Test Bank

Run:

```powershell
npm.cmd test
```

The test bank is in `data/test-bank/manifest.json`. It currently protects:

- Meshayu modern Tally service ledger
- Manthan legacy numeric Tally ledger
- Balaji multi-PDF statement ledger

When adding a new party format, add it to the test bank after it reconciles correctly. Future code changes should pass all old cases before being used by the account team.
