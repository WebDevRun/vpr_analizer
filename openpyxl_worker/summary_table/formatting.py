import logging

from openpyxl.formatting.rule import ColorScale, FormatObject, Rule
from openpyxl.styles import Alignment, Color
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_worker.constants import BRICK_COLOR, LIME_COLOR, THIN_BORDER, YELLOW_COLOR
from openpyxl_worker.types import FormatArgs, LineCells, MatrixCells


def format_point_cells(cells: LineCells, format_args: FormatArgs) -> None:
    """Apply alignment and number formatting to a set of cells.

    Args:
        cells (LineCells): The cells to format.
        format_args (FormatArgs): Formatting arguments (alignment, number format, wrap).
    """
    for cell in cells:
        cell.alignment = Alignment(
            horizontal=format_args.alignment.horizontal,
            vertical=format_args.alignment.vertical,
            wrap_text=format_args.wrap_text,
        )
        if format_args.number_format.value:
            cell.number_format = format_args.number_format.value


def set_borders(cells: MatrixCells) -> None:
    """Apply thin borders to all cells in the matrix.

    Args:
        cells (MatrixCells): Matrix of cells to apply borders to.
    """
    for row in cells:
        for cell in row:
            cell.border = THIN_BORDER


def generate_percentage_color_rule() -> Rule:
    """Generate a color scale rule for percentage formatting.

    Returns:
        Rule: An openpyxl Rule object for color scale formatting.
    """
    thresholds = [
        FormatObject(type="num", val=0),
        FormatObject(type="num", val=0.5),
        FormatObject(type="num", val=1),
    ]
    colors = [
        Color(BRICK_COLOR),
        Color(YELLOW_COLOR),
        Color(LIME_COLOR),
    ]
    color_scale = ColorScale(cfvo=thresholds, color=colors)
    return Rule(type="colorScale", colorScale=color_scale)


def apply_percentage_color_formatting(
    ws: Worksheet, start_coordinate: str, end_coordinate: str
) -> None:
    """Apply color scale conditional formatting to a range of percentage cells.

    Args:
        ws (Worksheet): The worksheet to apply formatting to.
        start_coordinate (str): Start cell coordinate (e.g., 'D2').
        end_coordinate (str): End cell coordinate (e.g., 'D10').
    """
    color_rule = generate_percentage_color_rule()
    ws.conditional_formatting.add(f"{start_coordinate}:{end_coordinate}", color_rule)
    logging.info(
        "Applied percentage color formatting to range: %s:%s",
        start_coordinate,
        end_coordinate,
    )
