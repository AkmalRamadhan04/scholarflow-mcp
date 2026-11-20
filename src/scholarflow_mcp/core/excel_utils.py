"""Shared Excel styling and Excel-safe DataValidation utilities."""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
OVERLAP_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
PASS_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
WARN_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

HEADER_FONT = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", size=10, bold=True)
LINK_FONT = Font(name="Calibri", size=10, color="0563C1", underline="single")

THIN_BORDER = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)


def style_sheet(
    ws: Worksheet,
    col_widths: dict[int, float] | None = None,
    freeze: str = "A2",
) -> None:
    """Apply consistent publication-grade table styling to a worksheet."""
    ws.freeze_panes = freeze
    max_row = max(ws.max_row, 1)
    max_col = max(ws.max_column, 1)
    for r_idx in range(1, max_row + 1):
        for c_idx in range(1, max_col + 1):
            cell = ws.cell(r_idx, c_idx)
            cell.border = THIN_BORDER
            if r_idx == 1:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                if cell.hyperlink is None and (cell.font is None or cell.font.color is None):
                    cell.font = BODY_FONT
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    if col_widths:
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width
    else:
        for c_idx in range(1, max_col + 1):
            col_letter = get_column_letter(c_idx)
            best = 14
            for r_idx in range(1, min(max_row, 60) + 1):
                val = ws.cell(r_idx, c_idx).value
                if val is not None:
                    best = max(best, min(len(str(val)) + 3, 45))
            ws.column_dimensions[col_letter].width = best


def add_safe_range_dropdown(
    ws: Worksheet,
    col_letter: str,
    start_row: int,
    end_row: int,
    ref_sheet_name: str,
    ref_col_letter: str,
    ref_start_row: int,
    ref_end_row: int,
    prompt_title: str = "Closed Codebook",
    prompt_text: str = "Select a value from the codebook dropdown.",
) -> DataValidation:
    """Add an Excel-safe cell-range DataValidation (avoids Excel's 255-char inline list corruption)."""
    safe_sheet = f"'{ref_sheet_name}'" if " " in ref_sheet_name else ref_sheet_name
    formula = f"={safe_sheet}!${ref_col_letter}${ref_start_row}:${ref_col_letter}${ref_end_row}"
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.promptTitle = prompt_title
    dv.prompt = prompt_text
    dv.errorTitle = "Invalid Code"
    dv.error = "Value must match one of the closed codebook entries."
    sqref = f"{col_letter}{start_row}:{col_letter}{max(end_row, start_row)}"
    dv.add(sqref)
    ws.add_data_validation(dv)
    return dv


def detect_unsafe_inline_validations(ws: Worksheet) -> list[str]:
    """Return any inline string DataValidation formulas exceeding Excel's 255-character limit."""
    issues: list[str] = []
    for dv in ws.data_validations.dataValidation:
        f1 = dv.formula1 or ""
        if f1.startswith('"') and len(f1) > 255:
            issues.append(f"{dv.sqref}: inline formula length {len(f1)} > 255")
    return issues
