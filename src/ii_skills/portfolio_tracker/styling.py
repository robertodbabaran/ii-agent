#!/usr/bin/env python3
"""
Bloomberg Terminal Styling Constants

Professional dark-themed styling for portfolio tracker workbooks.
Mirrors institutional Bloomberg Terminal aesthetics.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ── Color Palette ──────────────────────────────────────────────
NAVY = "1C2541"
DARK_BG = "0B132B"
LIGHT_ROW = "1C2541"
DARK_ROW = "141E33"
WHITE = "FFFFFF"
GREEN = "007A33"
RED = "B81D13"
AMBER = "FFB000"
LIGHT_GRAY = "D9D9D9"
MEDIUM_GRAY = "808080"
BLUE_INPUT = "0066CC"
BORDER_COLOR = "3A506B"

# ── Fonts ──────────────────────────────────────────────────────
HEADER_FONT = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
DATA_FONT = Font(name="Segoe UI", size=9, color=WHITE)
DATA_FONT_GREEN = Font(name="Segoe UI", size=9, color=GREEN)
DATA_FONT_RED = Font(name="Segoe UI", size=9, color=RED)
DATA_FONT_AMBER = Font(name="Segoe UI", size=9, color=AMBER)
TITLE_FONT = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
SUBTITLE_FONT = Font(name="Segoe UI", size=11, bold=True, color=LIGHT_GRAY)
KPI_VALUE_FONT = Font(name="Segoe UI", size=16, bold=True, color=WHITE)
KPI_LABEL_FONT = Font(name="Segoe UI", size=9, color=MEDIUM_GRAY)
INPUT_FONT = Font(name="Segoe UI", size=9, color=BLUE_INPUT)

# ── Fills ──────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
DARK_BG_FILL = PatternFill(start_color=DARK_BG, end_color=DARK_BG, fill_type="solid")
LIGHT_ROW_FILL = PatternFill(start_color=LIGHT_ROW, end_color=LIGHT_ROW, fill_type="solid")
DARK_ROW_FILL = PatternFill(start_color=DARK_ROW, end_color=DARK_ROW, fill_type="solid")
GREEN_FILL = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
RED_FILL = PatternFill(start_color=RED, end_color=RED, fill_type="solid")

# ── Alignment ──────────────────────────────────────────────────
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

# ── Borders ────────────────────────────────────────────────────
THIN_BORDER = Border(
    bottom=Side(style="thin", color=BORDER_COLOR),
)
HEADER_BORDER = Border(
    bottom=Side(style="medium", color=WHITE),
)

# ── Number Formats ─────────────────────────────────────────────
FMT_CURRENCY = '#,##0.00'
FMT_CURRENCY_K = '#,##0'
FMT_PCT = '0.0%'
FMT_PCT_2 = '0.00%'
FMT_NUMBER = '#,##0'
FMT_DATE = 'YYYY-MM-DD'
FMT_RATIO = '0.00x'
FMT_MULTIPLE = '0.0x'

# ── Sheet Settings ─────────────────────────────────────────────
ZOOM_SCALE = 85
FREEZE_CELL = "B2"


def apply_sheet_defaults(ws):
    """Apply Bloomberg Terminal defaults to a worksheet."""
    ws.sheet_view.zoomScale = ZOOM_SCALE
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = FREEZE_CELL
    ws.sheet_properties.tabColor = NAVY


def style_header_row(ws, row, max_col):
    """Style a header row with navy background and white text."""
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = HEADER_BORDER


def style_data_row(ws, row, max_col, zebra_index=0):
    """Style a data row with alternating fills."""
    fill = LIGHT_ROW_FILL if zebra_index % 2 == 0 else DARK_ROW_FILL
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = DATA_FONT
        cell.fill = fill
        cell.border = THIN_BORDER


def style_value_cell(ws, row, col, value, fmt=None):
    """Style a cell with conditional green/red coloring for numeric values."""
    cell = ws.cell(row=row, column=col)
    if fmt:
        cell.number_format = fmt
    if isinstance(value, (int, float)):
        if value > 0:
            cell.font = DATA_FONT_GREEN
        elif value < 0:
            cell.font = DATA_FONT_RED
    elif isinstance(value, str) and value.startswith("="):
        # For formulas, use default data font (Excel will compute the value)
        cell.font = DATA_FONT


def apply_bloomberg_styling(wb):
    """Apply Bloomberg Terminal styling across all sheets in workbook."""
    for ws in wb.worksheets:
        apply_sheet_defaults(ws)
        # Style first row as header if it has content
        max_col = ws.max_column or 1
        if ws.max_row and ws.max_row >= 1:
            style_header_row(ws, 1, max_col)
            for row_idx in range(2, (ws.max_row or 1) + 1):
                style_data_row(ws, row_idx, max_col, zebra_index=row_idx - 2)
