# Output Folder and Workbook Naming

Create one output folder per run with this exact format:

- `YYYY-MM-DD Costco Receipt`
- Example: `2026-02-07 Costco Receipt`

Create it in the current working directory unless user specifies another base path.

## Required output

Exactly one workbook file in that folder:

- `YYYY-MM-DD Costco Receipt.xlsx`

Workbook sheet names:

1. `Actual Values`
2. `Live Formula`
3. `Notes` (only if unresolved allocations exist)

## Minimal command pattern

```bash
DAY="$(date +%F)"
OUT_DIR="${DAY} Costco Receipt"
mkdir -p "$OUT_DIR"
OUT_XLSX="$OUT_DIR/${DAY} Costco Receipt.xlsx"
```
