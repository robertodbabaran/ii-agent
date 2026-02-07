# Output Contract

Use this exact column sequence in both tables:

1. Description
2. Unit Price
3. RJ Vol
4. Therese Vol
5. Mom Vol
6. Total Vol
7. RJ $
8. Therese $
9. Mom $
10. Total Line $

## Required summary rows (in this order)

- `SUBTOTAL`
- `TAX`
- `TOTAL`

## Formula conventions

- Use `=` formulas only in the Live Formula Table for computed cells.
- Use absolute references on subtotal/tax base cells where appropriate (for example, `$B$16`).
- Keep formulas simple and transparent (no named ranges required).

## Rounding

- Round displayed money values to 2 decimals.
- Keep volume values as provided by split instructions.
