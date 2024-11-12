from dataclasses import astuple
from typing import List

from openpyxl import Workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScale, FormatObject, Rule
from openpyxl.styles import Alignment, Color

from openpyxl_worker.constants import (
    BRICK_COLOR,
    LEFT_TOP_ALIGN,
    LIME_COLOR,
    THEME_RESULT_TABLE_HEADERS,
    THIN_BORDER,
    YELLOW_COLOR,
)
from openpyxl_worker.types import (
    FormatArgs,
    LineCells,
    MatrixCells,
    OverallResult,
    ResultCells,
    WorksheetRanges,
)


class SummaryTableWorker:
    NUMBER_COLUMN = 1
    TASK_NUMBER_COLUMN = 2
    THEME_COLUMN = 3
    POINT_COLUMN = 4

    def __init__(self, wb: Workbook, name: str):
        self.wb = wb

        if name not in self.wb.sheetnames:
            self.ws = self.wb.create_sheet(name)
            return

        self.ws = self.wb[name]

    def create(self, summary_table_data: List[WorksheetRanges]):
        self.add_header()
        result_cells = self.fill_table(summary_table_data)
        self.format_ws(result_cells)
        self.add_filter()

    def add_header(self):
        start_row = 1

        for str in THEME_RESULT_TABLE_HEADERS:
            column = THEME_RESULT_TABLE_HEADERS.index(str) + 1
            self.ws.cell(row=start_row, column=column, value=str)

    def fill_table(self, summary_table_data: List[WorksheetRanges]):
        start_row = 2
        ws_list: List[OverallResult] = []

        for wb in summary_table_data:
            for index, cell in enumerate(wb.task_cells):
                number = self.ws.cell(
                    row=start_row, column=self.NUMBER_COLUMN, value=start_row - 1
                )
                cell_value = f"='{wb.name}'!{cell.coordinate}"
                task_number = self.ws.cell(
                    row=start_row,
                    column=self.TASK_NUMBER_COLUMN,
                    value=cell_value,
                )
                cell_value = f"='{wb.name}'!B{cell.row}"
                task_name = self.ws.cell(
                    row=start_row,
                    column=self.THEME_COLUMN,
                    value=cell_value,
                )
                cell_value = f"='{wb.name}'!{wb.percentage_of_completion_formulas[index].coordinate}"
                percentage_of_completion = self.ws.cell(
                    row=start_row,
                    column=self.POINT_COLUMN,
                    value=cell_value,
                )
                start_row += 1
                overall_result = OverallResult(
                    number,
                    task_number,
                    task_name,
                    percentage_of_completion,
                )
                ws_list.append(overall_result)

        return ResultCells(self.ws.title, ws_list)

    def format_ws(self, result_cells: ResultCells):
        self.ws.conditional_formatting = ConditionalFormattingList()
        number_task_format = FormatArgs(LEFT_TOP_ALIGN)
        task_number_cells = tuple(
            [cells.task_number for cells in result_cells.overall_result]
        )
        cells = self.format_point_cells(task_number_cells, number_task_format)
        horizontal_cells = astuple(result_cells.overall_result[0])
        table_cells = self.find_table_cells(horizontal_cells, task_number_cells)
        self.set_borders(table_cells)
        percent_color_rule = self.generate_percentage_color_rule()
        start_coordinate = result_cells.overall_result[
            0
        ].percentage_of_completion.coordinate
        end_coordinate = result_cells.overall_result[
            -1
        ].percentage_of_completion.coordinate
        self.ws.conditional_formatting.add(
            f"{start_coordinate}:{end_coordinate}",
            percent_color_rule,
        )
        return cells

    def find_table_cells(
        self,
        horizontal_cells: LineCells,
        vertical_cells: LineCells,
    ):
        start_column = horizontal_cells[0].column
        end_column = horizontal_cells[-1].column
        start_row = horizontal_cells[0].row
        end_row = vertical_cells[-1].row

        return tuple(
            self.ws.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=start_column,
                max_col=end_column,
            )
        )

    def format_point_cells(self, cells: LineCells, format_args: FormatArgs):
        for cell in cells:
            cell.alignment = Alignment(
                format_args.alignment.horizontal,
                format_args.alignment.vertical,
                wrap_text=format_args.wrap_text,
            )

            if format_args.number_format.value:
                cell.number_format = format_args.number_format.value

        return cells

    def set_borders(self, cells: MatrixCells):
        for row in cells:
            for cell in row:
                cell.border = THIN_BORDER

        return cells

    def generate_percentage_color_rule(self):
        first = FormatObject(type="num", val=0)
        mid = FormatObject(type="num", val=0.5)
        last = FormatObject(type="num", val=1)
        colors = [
            Color(BRICK_COLOR),
            Color(YELLOW_COLOR),
            Color(LIME_COLOR),
        ]
        color_scale = ColorScale(cfvo=[first, mid, last], color=colors)
        return Rule(type="colorScale", colorScale=color_scale)

    def add_filter(self):
        self.ws.auto_filter.ref = self.ws.dimensions
