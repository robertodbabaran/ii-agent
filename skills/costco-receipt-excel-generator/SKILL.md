---
name: costco-receipt-excel-generator
description: Generate Costco receipt splits from receipt images into Excel output only. Use when users want a single dated output folder and one .xlsx workbook containing both Actual Values and Live Formula tables with editable formulas, per-person allocations, proportional tax, and totals.
---

# Costco Receipt Excel Generator

Generate Excel-only output in a single dated folder.

## Workflow

1. Extract every billable line from all provided receipt images.
2. Preserve quantity notation in `Description` when present (example: `4 @ 7.99`) and include discounts as separate negative lines.
3. Use this exact column order:
   - `Description | Unit Price | RJ Vol | Therese Vol | Mom Vol | Total Vol | RJ $ | Therese $ | Mom $ | Total Line $`
4. Apply split instructions to `RJ Vol`, `Therese Vol`, and `Mom Vol`.
5. Build one Excel workbook with two sheets:
   - `Actual Values`
   - `Live Formula`
6. Create output folder and workbook name per `references/output-files.md`.

## Calculation Rules

- `Total Vol = RJ Vol + Therese Vol + Mom Vol`.
- Item row amounts:
  - `RJ $ = Unit Price * RJ Vol`
  - `Therese $ = Unit Price * Therese Vol`
  - `Mom $ = Unit Price * Mom Vol`
  - `Total Line $ = RJ $ + Therese $ + Mom $`
- Include summary rows in this order: `SUBTOTAL`, `TAX`, `TOTAL`.
- Split tax proportionally by each person's subtotal share.

## Live Formula Sheet Rules

Assume header row is 1 and item rows start at 2.

- A Description
- B Unit Price
- C RJ Vol
- D Therese Vol
- E Mom Vol
- F Total Vol
- G RJ $
- H Therese $
- I Mom $
- J Total Line $

Per item row `r`:
- `F[r] = C[r]+D[r]+E[r]`
- `G[r] = B[r]*C[r]`
- `H[r] = B[r]*D[r]`
- `I[r] = B[r]*E[r]`
- `J[r] = G[r]+H[r]+I[r]`

For subtotal row `s` after last item row `n`:
- `B[s]` = receipt subtotal
- `G[s] = SUM(G2:G[n])`
- `H[s] = SUM(H2:H[n])`
- `I[s] = SUM(I2:I[n])`
- `J[s] = SUM(J2:J[n])`

For tax row `t=s+1`:
- `B[t]` = receipt tax
- `G[t] = (G[s]/$B$[s])*$B$[t]`
- `H[t] = (H[s]/$B$[s])*$B$[t]`
- `I[t] = (I[s]/$B$[s])*$B$[t]`
- `J[t] = SUM(G[t]:I[t])`

For total row `u=t+1`:
- `B[u] = B[s]+B[t]`
- `G[u] = G[s]+G[t]`
- `H[u] = H[s]+H[t]`
- `I[u] = I[s]+I[t]`
- `J[u] = J[s]+J[t]`

## Output Rules

- Do not generate markdown tables, CSV files, or TXT notes.
- Generate Excel output only.
- If there are unassigned split items, include an `UNASSIGNED` section in a third sheet named `Notes`; otherwise omit `Notes`.
- Provide a standalone prompt from `references/terminal-prompt-template.md` (no other skills mentioned).
