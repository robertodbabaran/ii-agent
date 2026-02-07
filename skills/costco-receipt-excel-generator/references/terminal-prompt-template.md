# Terminal Prompt Template (Excel Only)

```text
Task: From attached Costco receipt image(s), create Excel output only.

Create one folder in CWD named "$(date +%F) Costco Receipt".
Create one workbook in that folder named "YYYY-MM-DD Costco Receipt.xlsx" (same date as folder).

Workbook sheets:
1) Actual Values
2) Live Formula
3) Notes (only if unresolved allocations exist)

Use exact columns in both main sheets:
Description | Unit Price | RJ Vol | Therese Vol | Mom Vol | Total Vol | RJ $ | Therese $ | Mom $ | Total Line $

Rules:
- Extract all receipt item lines, including quantity notation and discounts.
- Total Vol=C+D+E.
- Item-row formulas in Live Formula sheet: G=B*C, H=B*D, I=B*E, J=G+H+I.
- Include SUBTOTAL, TAX, TOTAL rows.
- Subtotal rows use SUM formulas.
- TAX rows prorate by each person's subtotal share.
- TOTAL rows are subtotal+tax.
- Round displayed money to 2 decimals.

Splits:
[PASTE SPLITS]

Output constraints:
- Do not output markdown tables, CSV, or TXT files.
- Excel workbook is the only output artifact.
- Print absolute workbook path when complete.
```
