from typing import Tuple

from openpyxl.cell.cell import Cell
from openpyxl.formatting.rule import ColorScale, FormatObject, Rule
from openpyxl.styles import Alignment, Color
from openpyxl.styles.borders import BORDER_THIN, Border, Side

from openpyxl_worker.cells_finder import MatrixCells
from openpyxl_worker.types import AlignmentCell, FormatArgs


class CellFormater:
    LEFT_TOP_ALIGN = AlignmentCell("left", "top")
    RIGHT_TOP_ALIGN = AlignmentCell("right", "top")
    LIME_COLOR = "81d41a"
    YELLOW_COLOR = "ffff00"
    BRICK_COLOR = "ff4000"

    def __init__(self):
        self.thin_border = Border(
            left=Side(style=BORDER_THIN),
            right=Side(style=BORDER_THIN),
            top=Side(style=BORDER_THIN),
            bottom=Side(style=BORDER_THIN),
        )

    def format_not_point_cells(self, cells: Tuple[Cell, ...], format_args: FormatArgs):
        for cell in cells:
            cell.alignment = Alignment(
                format_args.alignment.horizontal,
                format_args.alignment.vertical,
                wrap_text=format_args.wrap_text,
            )

            if format_args.number_format.value:
                cell.number_format = format_args.number_format.value

        return cells

    def format_point_cells(self, cells: MatrixCells, format_args: FormatArgs):
        for row in cells:
            self.format_not_point_cells(row, format_args)

        return cells

    def set_borders(self, cells: MatrixCells):
        for row in cells:
            for cell in row:
                cell.border = self.thin_border

        return cells

    def generate_percentage_color_rule(self):
        first = FormatObject(type="num", val=0)
        mid = FormatObject(type="num", val=0.5)
        last = FormatObject(type="num", val=1)
        colors = [
            Color(self.BRICK_COLOR),
            Color(self.YELLOW_COLOR),
            Color(self.LIME_COLOR),
        ]
        color_scale = ColorScale(cfvo=[first, mid, last], color=colors)
        return Rule(type="colorScale", colorScale=color_scale)

    def generate_point_color_rule(self, max_point: int):
        first = FormatObject(type="num", val=0)
        mid = FormatObject(type="num", val=max_point / 2)
        last = FormatObject(type="num", val=max_point)
        colors = [
            Color(self.BRICK_COLOR),
            Color(self.YELLOW_COLOR),
            Color(self.LIME_COLOR),
        ]
        color_scale = ColorScale(cfvo=[first, mid, last], color=colors)
        return Rule(type="colorScale", colorScale=color_scale)
