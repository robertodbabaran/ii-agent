# Output Folder and File Naming

Always create a dedicated output folder with this exact name format:

- `YYYY-MM-DD Costco Receipt`
- Example: `2026-02-07 Costco Receipt`

Create it in the current working directory unless the user gives a different base path.

## Required files in that folder

1. `actual-values-table.md` - markdown table for human review
2. `live-formula-table.md` - markdown version of formula table
3. `live-formula-table.csv` - Excel paste-ready CSV with formulas
4. `notes.txt` - unresolved split notes (write `UNASSIGNED: none` if fully assigned)

## Minimal command pattern

```bash
OUT_DIR="$(date +%F) Costco Receipt"
mkdir -p "$OUT_DIR"
```
