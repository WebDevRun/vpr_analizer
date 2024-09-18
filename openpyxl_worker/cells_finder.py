from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_worker.constants import THEME_TABLE_HEADERS
from openpyxl_worker.types import LineCells, MatrixCells, Range


class CellsFinder:
    def __init__(self, ws: Worksheet) -> None:
        self.ws = ws

    def getCells(self, range: Range) -> MatrixCells:
        return self.ws[range.start : range.end]

    def find_students(self, point_cells: MatrixCells) -> LineCells:
        start_row = point_cells[0][0].row
        end_row = point_cells[-1][0].row
        range = Range(f"A{start_row}", f"A{end_row}")
        cells = self.getCells(range)
        student_cells = [row[0] for row in cells]
        return tuple(student_cells)

    def find_task_cells(self, point_cells: MatrixCells) -> LineCells:
        tasks_row = point_cells[0][0].row - 1
        start_column_letter = point_cells[0][0].column_letter
        end_column_letter = point_cells[0][-1].column_letter
        start_coordinate = f"{start_column_letter}{tasks_row}"
        end_coordinate = f"{end_column_letter}{tasks_row}"
        task_range = Range(start_coordinate, end_coordinate)
        return self.getCells(task_range)[0]

    def find_student_formulas(self, header_cells: LineCells) -> LineCells:
        student_cells = [
            cell for cell in header_cells if cell.value not in THEME_TABLE_HEADERS
        ]
        return tuple(student_cells)

    def find_table_cells(
        self,
        horizontal_cells: LineCells,
        vertical_cells: LineCells,
    ):
        start_column = horizontal_cells[0].column_letter
        end_column = horizontal_cells[-1].column_letter
        start_row = horizontal_cells[0].row
        end_row = vertical_cells[-1].row
        table_range = Range(
            f"{start_column}{start_row}",
            f"{end_column}{end_row}",
        )
        return self.getCells(table_range)

    def find_max_point_cells(
        self, number_formulas: LineCells, table_header: Cell
    ) -> LineCells:
        row = table_header.row + 1
        column = table_header.column
        start_cell = self.ws.cell(row, column)
        end_cell = self.ws.cell(number_formulas[-1].row, column)
        range = Range(start_cell.coordinate, end_cell.coordinate)
        cells = [cell[0] for cell in self.getCells(range)]

        return tuple(cells)
