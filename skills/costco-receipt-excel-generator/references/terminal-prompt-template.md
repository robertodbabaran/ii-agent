# Terminal Prompt Templates

Use one of the templates below.

## A) Ultra-Compact Prompt (lowest tokens)

```text
Task: From attached Costco receipt image(s), create split outputs only.

Create folder: "$(date +%F) Costco Receipt" in CWD.
Write files in that folder:
- actual-values-table.md
- live-formula-table.md
- live-formula-table.csv
- notes.txt

Columns (exact order):
Description | Unit Price | RJ Vol | Therese Vol | Mom Vol | Total Vol | RJ $ | Therese $ | Mom $ | Total Line $

Rules:
- Extract all item lines, including qty notation and discounts.
- Total Vol=C+D+E.
- Formula table item rows: G=B*C, H=B*D, I=B*E, J=G+H+I.
- Include SUBTOTAL, TAX, TOTAL rows.
- Subtotal rows use SUM.
- TAX rows prorate by person subtotal share.
- TOTAL rows are subtotal+tax.
- Money rounded to 2 decimals.

Splits:
[PASTE SPLITS]

Outputs:
- actual-values-table.md: Actual Values markdown table
- live-formula-table.md: Live Formula markdown table
- live-formula-table.csv: same Live Formula table as CSV with formulas
- notes.txt: unresolved items as UNASSIGNED lines, else "UNASSIGNED: none"

Finally print absolute output folder path and file list.
```

## B) Explicit Prompt (more verbose, easier to audit)

```text
Analyze the attached Costco receipt image(s) and do only this task.

Create a folder named "YYYY-MM-DD Costco Receipt" (today's date) in the current working directory.

Inside that folder, write exactly these files:
1) actual-values-table.md
2) live-formula-table.md
3) live-formula-table.csv
4) notes.txt

Generate two tables using this exact column order:
Description | Unit Price | RJ Vol | Therese Vol | Mom Vol | Total Vol | RJ $ | Therese $ | Mom $ | Total Line $

Rules:
- Extract every line item from the receipt, including quantity notation (example: 4 @ 7.99) and discounts.
- Keep Unit Price as listed on receipt.
- Total Vol = RJ Vol + Therese Vol + Mom Vol.
- For item rows in formula table:
  - RJ $ = B*C
  - Therese $ = B*D
  - Mom $ = B*E
  - Total Line $ = G+H+I
- Include SUBTOTAL, TAX, TOTAL rows.
- For formula table subtotal rows use SUM formulas.
- For TAX row, distribute tax proportionally from each person's subtotal share.
- For TOTAL row, subtotal + tax.
- Round displayed money to 2 decimals.

Split allocations for this run:
[PASTE SPLIT ALLOCATIONS HERE]

Output requirements:
- Put Actual Values markdown table in actual-values-table.md.
- Put Live Formula markdown table in live-formula-table.md.
- Put Live Formula CSV (with formulas) in live-formula-table.csv.
- Put unresolved allocations in notes.txt as UNASSIGNED entries, or "UNASSIGNED: none".
- In terminal output, print the absolute path of the created folder and list the files.
```
