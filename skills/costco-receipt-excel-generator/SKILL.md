---
name: costco-receipt-excel-generator
description: Generate itemized Costco receipt splits from one or more receipt images into two outputs: (1) an Actual Values Table for review and (2) a Live Formula Table that can be pasted into Excel with editable formulas for per-person allocation, subtotal, proportional tax, and totals. Use when users ask to split Costco (or similar warehouse) receipts across people, include quantity-based lines/discounts, produce spreadsheet-ready formula rows, and save results into a dated receipt output folder.
---

# Costco Receipt Excel Generator

Produce two tables from receipt images and split instructions:
1. **Actual Values Table** (numerical values, reviewer-friendly)
2. **Live Formula Table** (Excel formulas, paste-ready)

## Workflow

1. Extract every billable line from all provided receipt images.
2. Preserve quantity notation in the description when present (example: `4 @ 7.99`) and include discounts as separate negative lines.
3. Normalize to these columns in this exact order:
   - `Description | Unit Price | RJ Vol | Therese Vol | Mom Vol | Total Vol | RJ $ | Therese $ | Mom $ | Total Line $`
4. Apply user split instructions to volume columns (`RJ Vol`, `Therese Vol`, `Mom Vol`).
5. Compute `Total Vol` per line as `RJ Vol + Therese Vol + Mom Vol`.
6. Build both tables using the output contract in `references/output-contract.md`.
7. Create a dedicated output folder and write files using the naming convention in `references/output-files.md`.

## Calculation Rules

- Keep `Unit Price` as receipt unit price, not extended price.
- For regular item lines in **Actual Values Table**:
  - `RJ $ = Unit Price * RJ Vol`
  - `Therese $ = Unit Price * Therese Vol`
  - `Mom $ = Unit Price * Mom Vol`
  - `Total Line $ = RJ $ + Therese $ + Mom $`
- Include subtotal/tax/total summary rows at the bottom.
- Split tax proportionally by each person's subtotal share.

## Live Formula Table Rules

Assume headers are in row 1 and data starts in row 2.

- Columns:
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
- For each item row `r`:
  - `F[r] = C[r]+D[r]+E[r]`
  - `G[r] = B[r]*C[r]`
  - `H[r] = B[r]*D[r]`
  - `I[r] = B[r]*E[r]`
  - `J[r] = G[r]+H[r]+I[r]`
- For subtotal row `s` (after last item row `n`):
  - `B[s]` = receipt subtotal value
  - `G[s] = SUM(G2:G[n])`
  - `H[s] = SUM(H2:H[n])`
  - `I[s] = SUM(I2:I[n])`
  - `J[s] = SUM(J2:J[n])`
- For tax row `t = s+1`:
  - `B[t]` = receipt tax value
  - `G[t] = (G[s]/$B$[s])*$B$[t]`
  - `H[t] = (H[s]/$B$[s])*$B$[t]`
  - `I[t] = (I[s]/$B$[s])*$B$[t]`
  - `J[t] = SUM(G[t]:I[t])`
- For total row `u = t+1`:
  - `B[u] = B[s]+B[t]`
  - `G[u] = G[s]+G[t]`
  - `H[u] = H[s]+H[t]`
  - `I[u] = I[s]+I[t]`
  - `J[u] = J[s]+J[t]`

## Quality Checks

Before finalizing output:
- Ensure every item line from the receipt appears exactly once (plus discount lines).
- Ensure each line's `Total Vol` matches sum of person volumes.
- Ensure sum of person subtotals equals receipt subtotal (allow small rounding drift of ±0.01).
- Ensure sum of person taxes equals receipt tax (allow ±0.01).
- Ensure grand total equals subtotal + tax.

## Output Style

- Return both tables as markdown tables.
- Immediately after the markdown tables, return the **Live Formula Table** again as CSV text so users can paste into Excel directly.
- If split instructions are missing for any line, add an `UNASSIGNED` note after the tables listing unresolved items.
- Provide a **standalone terminal prompt** from `references/terminal-prompt-template.md` and prefer the **Ultra-Compact Prompt** to minimize token usage (do not mention other skills).
