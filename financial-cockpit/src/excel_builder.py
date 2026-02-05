"""
Excel Builder Utilities for Financiële Cockpit
Contains styling, table creation, and validation functions.
"""

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment,
    NamedStyle,
    Protection,
)
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional, Tuple

from model_config import COLORS, COLUMN_WIDTHS, DATA_VALIDATIONS


# =============================================================================
# STYLE DEFINITIONS
# =============================================================================

def create_named_styles(wb: Workbook) -> None:
    """Create and register named styles in the workbook."""

    # Header style
    header_style = NamedStyle(name="header_style")
    header_style.font = Font(bold=True, color=COLORS["header_font"], size=11)
    header_style.fill = PatternFill(
        start_color=COLORS["header_bg"], end_color=COLORS["header_bg"], fill_type="solid"
    )
    header_style.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_style.border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Title style (for dashboard headers)
    title_style = NamedStyle(name="title_style")
    title_style.font = Font(bold=True, size=14, color=COLORS["header_bg"])
    title_style.alignment = Alignment(horizontal="left", vertical="center")

    # Section header style
    section_style = NamedStyle(name="section_style")
    section_style.font = Font(bold=True, size=12, color=COLORS["header_bg"])
    section_style.fill = PatternFill(
        start_color=COLORS["neutral_bg"], end_color=COLORS["neutral_bg"], fill_type="solid"
    )
    section_style.alignment = Alignment(horizontal="left", vertical="center")

    # Currency style
    currency_style = NamedStyle(name="currency_style")
    currency_style.number_format = '€ #,##0.00'
    currency_style.alignment = Alignment(horizontal="right")

    # Percentage style
    percentage_style = NamedStyle(name="percentage_style")
    percentage_style.number_format = '0.00%'
    percentage_style.alignment = Alignment(horizontal="right")

    # Date style
    date_style = NamedStyle(name="date_style")
    date_style.number_format = 'YYYY-MM-DD'
    date_style.alignment = Alignment(horizontal="center")

    # Input cell style (for user-editable cells)
    input_style = NamedStyle(name="input_style")
    input_style.fill = PatternFill(
        start_color=COLORS["input_bg"], end_color=COLORS["input_bg"], fill_type="solid"
    )
    input_style.border = Border(
        left=Side(style="thin", color="999999"),
        right=Side(style="thin", color="999999"),
        top=Side(style="thin", color="999999"),
        bottom=Side(style="thin", color="999999"),
    )

    # Calculated cell style
    calc_style = NamedStyle(name="calc_style")
    calc_style.fill = PatternFill(
        start_color=COLORS["calc_bg"], end_color=COLORS["calc_bg"], fill_type="solid"
    )
    calc_style.protection = Protection(locked=True)

    # Status OK style
    ok_style = NamedStyle(name="ok_style")
    ok_style.font = Font(bold=True, color=COLORS["ok_font"])
    ok_style.fill = PatternFill(
        start_color=COLORS["ok_bg"], end_color=COLORS["ok_bg"], fill_type="solid"
    )
    ok_style.alignment = Alignment(horizontal="center")

    # Status ALERT style
    alert_style = NamedStyle(name="alert_style")
    alert_style.font = Font(bold=True, color=COLORS["alert_font"])
    alert_style.fill = PatternFill(
        start_color=COLORS["alert_bg"], end_color=COLORS["alert_bg"], fill_type="solid"
    )
    alert_style.alignment = Alignment(horizontal="center")

    # Status RISK style
    risk_style = NamedStyle(name="risk_style")
    risk_style.font = Font(bold=True, color=COLORS["risk_font"])
    risk_style.fill = PatternFill(
        start_color=COLORS["risk_bg"], end_color=COLORS["risk_bg"], fill_type="solid"
    )
    risk_style.alignment = Alignment(horizontal="center")

    # Register all styles
    styles = [
        header_style, title_style, section_style, currency_style,
        percentage_style, date_style, input_style, calc_style,
        ok_style, alert_style, risk_style
    ]

    for style in styles:
        try:
            wb.add_named_style(style)
        except ValueError:
            pass  # Style already exists


# =============================================================================
# TABLE CREATION
# =============================================================================

def create_excel_table(
    ws: Worksheet,
    table_name: str,
    columns: List[str],
    start_row: int = 2,
    start_col: int = 1,
    data: Optional[List[Dict[str, Any]]] = None,
    num_empty_rows: int = 5,
) -> Table:
    """
    Create an Excel table with headers and optional data.

    Args:
        ws: Worksheet to add table to
        table_name: Name of the table (e.g., 'tblBTC')
        columns: List of column headers
        start_row: Starting row (1-indexed)
        start_col: Starting column (1-indexed)
        data: Optional list of dictionaries with data
        num_empty_rows: Number of empty rows to add if no data

    Returns:
        The created Table object
    """
    # Write headers
    for col_idx, header in enumerate(columns, start=start_col):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.style = "header_style"

    # Determine number of data rows
    if data:
        num_rows = len(data)
    else:
        num_rows = num_empty_rows

    # Write data if provided
    if data:
        for row_idx, row_data in enumerate(data, start=start_row + 1):
            for col_idx, col_name in enumerate(columns, start=start_col):
                value = row_data.get(col_name, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)

    # Calculate table range
    end_col = start_col + len(columns) - 1
    end_row = start_row + num_rows

    start_cell = f"{get_column_letter(start_col)}{start_row}"
    end_cell = f"{get_column_letter(end_col)}{end_row}"
    table_range = f"{start_cell}:{end_cell}"

    # Create table
    table = Table(displayName=table_name, ref=table_range)

    # Apply table style
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style

    ws.add_table(table)

    return table


def add_table_formulas(
    ws: Worksheet,
    table_name: str,
    column_formulas: Dict[str, str],
    start_row: int,
    columns: List[str],
    num_rows: int,
    start_col: int = 1,
) -> None:
    """
    Add formulas to specific columns in a table.

    Args:
        ws: Worksheet
        table_name: Name of the table
        column_formulas: Dict mapping column name to formula template
        start_row: Data start row (header row + 1)
        columns: List of column names
        num_rows: Number of data rows
        start_col: Starting column
    """
    for col_name, formula in column_formulas.items():
        if col_name in columns:
            col_idx = columns.index(col_name) + start_col
            for row_idx in range(start_row, start_row + num_rows):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = formula


# =============================================================================
# DATA VALIDATION
# =============================================================================

def add_dropdown_validation(
    ws: Worksheet,
    validation_type: str,
    cell_range: str,
) -> None:
    """
    Add dropdown data validation to a cell range.

    Args:
        ws: Worksheet
        validation_type: Key from DATA_VALIDATIONS dict
        cell_range: Excel range string (e.g., 'D3:D100')
    """
    if validation_type not in DATA_VALIDATIONS:
        raise ValueError(f"Unknown validation type: {validation_type}")

    options = DATA_VALIDATIONS[validation_type]
    formula = '"' + ",".join(options) + '"'

    dv = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Ongeldige invoer",
        error=f"Kies een waarde uit de lijst: {', '.join(options)}",
    )

    ws.add_data_validation(dv)
    dv.add(cell_range)


# =============================================================================
# CONDITIONAL FORMATTING
# =============================================================================

def add_status_conditional_formatting(
    ws: Worksheet,
    cell_range: str,
) -> None:
    """
    Add conditional formatting for status columns (OK/ALERT/RISK).

    Args:
        ws: Worksheet
        cell_range: Excel range string (e.g., 'J3:J100')
    """
    # OK formatting
    ok_fill = PatternFill(
        start_color=COLORS["ok_bg"], end_color=COLORS["ok_bg"], fill_type="solid"
    )
    ok_font = Font(color=COLORS["ok_font"], bold=True)

    # ALERT formatting
    alert_fill = PatternFill(
        start_color=COLORS["alert_bg"], end_color=COLORS["alert_bg"], fill_type="solid"
    )
    alert_font = Font(color=COLORS["alert_font"], bold=True)

    # RISK formatting
    risk_fill = PatternFill(
        start_color=COLORS["risk_bg"], end_color=COLORS["risk_bg"], fill_type="solid"
    )
    risk_font = Font(color=COLORS["risk_font"], bold=True)

    # Add rules
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"OK"'], fill=ok_fill, font=ok_font)
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"ALERT"'], fill=alert_fill, font=alert_font)
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"RISK"'], fill=risk_fill, font=risk_font)
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"CRITICAL"'], fill=risk_fill, font=risk_font)
    )


def add_severity_conditional_formatting(
    ws: Worksheet,
    cell_range: str,
) -> None:
    """
    Add conditional formatting for severity columns (INFO/ALERT/RISK/CRITICAL).
    """
    # INFO formatting (neutral)
    info_fill = PatternFill(
        start_color=COLORS["neutral_bg"], end_color=COLORS["neutral_bg"], fill_type="solid"
    )
    info_font = Font(color=COLORS["header_bg"])

    # Reuse status formatting for ALERT/RISK/CRITICAL
    add_status_conditional_formatting(ws, cell_range)

    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"INFO"'], fill=info_fill, font=info_font)
    )


# =============================================================================
# COLUMN FORMATTING
# =============================================================================

def set_column_widths(
    ws: Worksheet,
    column_widths: Dict[int, float],
) -> None:
    """
    Set column widths for a worksheet.

    Args:
        ws: Worksheet
        column_widths: Dict mapping column index (1-based) to width
    """
    for col_idx, width in column_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width


def auto_column_widths(
    ws: Worksheet,
    columns: List[str],
    start_col: int = 1,
    min_width: float = 10,
    max_width: float = 50,
) -> None:
    """
    Auto-size columns based on header length with min/max constraints.
    """
    for idx, col_name in enumerate(columns, start=start_col):
        col_letter = get_column_letter(idx)
        # Calculate width based on header + some padding
        width = min(max(len(col_name) + 4, min_width), max_width)
        ws.column_dimensions[col_letter].width = width


def apply_number_formats(
    ws: Worksheet,
    column_formats: Dict[int, str],
    start_row: int,
    end_row: int,
) -> None:
    """
    Apply number formats to columns.

    Args:
        ws: Worksheet
        column_formats: Dict mapping column index to format string
        start_row: First data row
        end_row: Last data row
    """
    for col_idx, format_str in column_formats.items():
        for row_idx in range(start_row, end_row + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = format_str


# =============================================================================
# SHEET UTILITIES
# =============================================================================

def freeze_panes(ws: Worksheet, row: int = 2, col: int = 1) -> None:
    """Freeze panes at specified position (freeze above and to the left)."""
    ws.freeze_panes = ws.cell(row=row, column=col)


def add_sheet_title(
    ws: Worksheet,
    title: str,
    row: int = 1,
    col: int = 1,
) -> None:
    """Add a title to a sheet."""
    cell = ws.cell(row=row, column=col, value=title)
    cell.style = "title_style"


def add_section_header(
    ws: Worksheet,
    title: str,
    row: int,
    col: int = 1,
    span: int = 1,
) -> None:
    """Add a section header."""
    cell = ws.cell(row=row, column=col, value=title)
    cell.style = "section_style"
    if span > 1:
        ws.merge_cells(
            start_row=row, start_column=col,
            end_row=row, end_column=col + span - 1
        )


def add_instruction_text(
    ws: Worksheet,
    text: str,
    start_row: int,
    col: int = 1,
    wrap_width: int = 80,
) -> int:
    """
    Add instruction text to a sheet, wrapping at specified width.
    Returns the last row used.
    """
    lines = text.split('\n')
    current_row = start_row

    for line in lines:
        cell = ws.cell(row=current_row, column=col, value=line)
        cell.alignment = Alignment(wrap_text=True)
        current_row += 1

    return current_row - 1


# =============================================================================
# DASHBOARD HELPERS
# =============================================================================

def create_kpi_cell(
    ws: Worksheet,
    row: int,
    label_col: int,
    value_col: int,
    label: str,
    formula: str,
    number_format: str = '€ #,##0.00',
) -> None:
    """Create a KPI display with label and value."""
    label_cell = ws.cell(row=row, column=label_col, value=label)
    label_cell.font = Font(bold=True)
    label_cell.alignment = Alignment(horizontal="right")

    value_cell = ws.cell(row=row, column=value_col, value=formula)
    value_cell.number_format = number_format
    value_cell.font = Font(bold=True, size=12)
    value_cell.alignment = Alignment(horizontal="left")


def create_dashboard_section(
    ws: Worksheet,
    title: str,
    start_row: int,
    start_col: int,
    kpis: List[Dict[str, Any]],
) -> int:
    """
    Create a dashboard section with title and KPIs.
    Returns the last row used.
    """
    # Section title
    add_section_header(ws, title, start_row, start_col, span=3)

    current_row = start_row + 1
    for kpi in kpis:
        create_kpi_cell(
            ws,
            current_row,
            start_col,
            start_col + 1,
            kpi["label"],
            kpi["formula"],
            kpi.get("format", '€ #,##0.00'),
        )
        current_row += 1

    return current_row


# =============================================================================
# FORMULA HELPERS
# =============================================================================

def build_xlookup(
    lookup_value: str,
    lookup_array: str,
    return_array: str,
    if_not_found: str = '""',
) -> str:
    """Build an XLOOKUP formula string."""
    return f'=XLOOKUP({lookup_value},{lookup_array},{return_array},{if_not_found})'


def build_sumif(
    range_ref: str,
    criteria: str,
    sum_range: str,
) -> str:
    """Build a SUMIF formula string."""
    return f'=SUMIF({range_ref},{criteria},{sum_range})'


def build_if_formula(
    condition: str,
    true_value: str,
    false_value: str,
) -> str:
    """Build an IF formula string."""
    return f'=IF({condition},{true_value},{false_value})'


def build_nested_if(
    conditions: List[Tuple[str, str]],
    else_value: str,
) -> str:
    """
    Build a nested IF formula.

    Args:
        conditions: List of (condition, value) tuples
        else_value: Value if no conditions match
    """
    if not conditions:
        return else_value

    condition, value = conditions[0]
    remaining = conditions[1:]

    if remaining:
        nested = build_nested_if(remaining, else_value)
        return f'IF({condition},{value},{nested})'
    else:
        return f'IF({condition},{value},{else_value})'
