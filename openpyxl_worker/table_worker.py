from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList

from openpyxl_worker.cell_formater import CellFormater
from openpyxl_worker.cells_filler import CellsFiller
from openpyxl_worker.cells_finder import CellsFinder
from openpyxl_worker.constants import REPLACE_VALUES
from openpyxl_worker.types import FormatArgs, NumberFormatCell, Range, WorksheetRanges


class TableWorker:
    def __init__(self, file_path: Path):
        self.path = file_path
        self.wb = load_workbook(self.path)
        self.ws = self.wb[self.wb.sheetnames[0]]
        self.cells_finder = CellsFinder(self.ws)
        self.cells_filler = CellsFiller(self.ws)
        self.cells_formater = CellFormater()

    def activate_sheet(self, sheet_name: str):
        self.ws = self.wb[sheet_name]
        self.cells_finder = CellsFinder(self.ws)
        self.cells_filler = CellsFiller(self.ws)
        return self

    def replace_x_cells(self, cell_range: Range):
        selected_cell_range = self.cells_finder.getCells(cell_range)

        for cell_row in selected_cell_range:
            for cell in cell_row:
                if cell.value in REPLACE_VALUES:
                    cell.value = 0

        return self

    def fill_theme_table(self, point_range: Range):
        worksheet_ranges = self.calculate_values(point_range)
        self.format_worksheet(worksheet_ranges)

        return worksheet_ranges

    def calculate_values(self, point_range: Range):
        cells_finder = self.cells_finder
        cells_filler = self.cells_filler
        point_cells = cells_finder.getCells(point_range)
        student_cells = cells_finder.find_students(point_cells)
        table_headers = cells_filler.fill_table_header(student_cells)
        task_cells = cells_finder.find_task_cells(point_cells)
        task_formulas = cells_filler.fill_task_numbers(task_cells, table_headers[0])
        student_formulas = cells_finder.find_student_formulas(table_headers)
        point_formulas = cells_filler.fill_point_formulas(point_cells, student_formulas)
        average_formulas = cells_filler.fill_average_formulas(
            point_formulas, table_headers[-2]
        )
        percentage_of_completion_formulas = cells_filler.fill_percentage_of_completion(
            average_formulas,
            table_headers[2],
            table_headers[-1],
        )
        max_point_cells = cells_finder.find_max_point_cells(
            task_formulas, table_headers[2]
        )
        sum_max_point_formula = cells_filler.fill_sum_max_points(max_point_cells)
        sum_student_point_formulas = cells_filler.fill_sum_student_points(
            point_formulas
        )
        average_point = cells_filler.fill_average_point(average_formulas)
        average_percentage_of_completion = (
            cells_filler.fill_average_percentage_of_completion(
                sum_student_point_formulas,
                sum_max_point_formula,
                percentage_of_completion_formulas[-1],
            )
        )
        percentage_of_points = cells_filler.fill_percentage_of_point(
            sum_student_point_formulas, sum_max_point_formula
        )

        return WorksheetRanges(
            self.ws.title,
            table_headers,
            task_formulas,
            student_formulas,
            point_formulas,
            average_formulas,
            percentage_of_completion_formulas,
            max_point_cells,
            sum_max_point_formula,
            sum_student_point_formulas,
            average_point,
            average_percentage_of_completion,
            percentage_of_points,
        )

    def format_worksheet(self, worksheet_ranges: WorksheetRanges):
        cell_formater = self.cells_formater
        table_header_format = FormatArgs(cell_formater.LEFT_TOP_ALIGN, wrap_text=True)
        number_formula_format = FormatArgs(cell_formater.LEFT_TOP_ALIGN)
        point_formula_format = FormatArgs(cell_formater.RIGHT_TOP_ALIGN)
        average_formula_format = FormatArgs(
            cell_formater.RIGHT_TOP_ALIGN,
            NumberFormatCell.FORMAT_NUMBER_00,
        )
        percentage_formula_format = FormatArgs(
            cell_formater.RIGHT_TOP_ALIGN, NumberFormatCell.FORMAT_PERCENTAGE_00
        )
        cell_formater.format_not_point_cells(
            worksheet_ranges.table_headers, table_header_format
        )
        cell_formater.format_not_point_cells(
            worksheet_ranges.task_formulas, number_formula_format
        )
        cell_formater.format_point_cells(
            worksheet_ranges.point_formulas, point_formula_format
        )
        cell_formater.format_not_point_cells(
            worksheet_ranges.average_formulas, average_formula_format
        )
        cell_formater.format_not_point_cells(
            worksheet_ranges.percentage_of_completion_formulas,
            percentage_formula_format,
        )
        cell_formater.format_not_point_cells(
            (worksheet_ranges.average_point,), average_formula_format
        )
        cell_formater.format_not_point_cells(
            (worksheet_ranges.average_percentage_of_completion,),
            percentage_formula_format,
        )
        cell_formater.format_not_point_cells(
            worksheet_ranges.percentage_of_points, percentage_formula_format
        )
        table_cells = self.cells_finder.find_table_cells(
            worksheet_ranges.table_headers, worksheet_ranges.percentage_of_points
        )
        cell_formater.set_borders(table_cells)

    def paint_worksheet(self, worksheet_ranges: WorksheetRanges):
        cell_formater = self.cells_formater
        self.ws.conditional_formatting = ConditionalFormattingList()
        percent_color_rule = cell_formater.generate_percentage_color_rule()
        start_coordinate = worksheet_ranges.percentage_of_points[0].coordinate
        end_coordinate = worksheet_ranges.percentage_of_points[-1].coordinate
        self.ws.conditional_formatting.add(
            f"{start_coordinate}:{end_coordinate}",
            percent_color_rule,
        )
        start_coordinate = worksheet_ranges.percentage_of_completion_formulas[
            0
        ].coordinate
        end_coordinate = worksheet_ranges.average_percentage_of_completion.coordinate
        self.ws.conditional_formatting.add(
            f"{start_coordinate}:{end_coordinate}",
            percent_color_rule,
        )

        for index, cell in enumerate(worksheet_ranges.max_point_cells):
            if type(cell.value) is int:
                point_color_rule = cell_formater.generate_point_color_rule(cell.value)
            row = cell.row
            start_column = cell.column + 1
            end_column = worksheet_ranges.average_formulas[index].column - 1
            start_cell = self.ws.cell(row, start_column)
            end_cell = self.ws.cell(row, end_column)
            self.ws.conditional_formatting.add(
                f"{start_cell.coordinate}:{end_cell.coordinate}",
                point_color_rule,
            )

    def save_table(self, file_path: Path):
        self.wb.save(file_path)
