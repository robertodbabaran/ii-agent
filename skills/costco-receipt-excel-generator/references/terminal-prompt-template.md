# Terminal Prompt Template (Standalone / Low Token)

Use this exact prompt in your terminal agent call. Replace bracketed placeholders.

```text
Analyze the attached Costco receipt image(s) and do only this task.

Create a folder named "YYYY-MM-DD Costco Receipt" (today's date, format like 2026-02-07 Costco Receipt) in the current working directory.

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
  - RJ $ = B* C
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
