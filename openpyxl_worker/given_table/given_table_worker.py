from typing import List, Tuple

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_worker.given_table.constants import EMPTY_STUDENT, REPLACE_VALUES
from openpyxl_worker.types import (
    FilledRows,
    FinderCells,
    GivenTableCells,
    LineCells,
    MatrixCells,
    Range,
    TaskValues,
)


class GivenTableWorker:
    def __init__(self, ws: Worksheet, point_range: Range) -> None:
        self.ws = ws
        self.point_range = point_range

    def get_cell_ranges(self) -> GivenTableCells:
        filled_rows = self.select_filled_rows(self.point_range)
        point_cells = self.replace_x_cells(filled_rows.rows)
        cells = self.cell_range_finder(filled_rows.rows)
        task_values = self.select_task_values(cells.task_cells)
        return GivenTableCells(
            point_cells,
            cells.student_cells,
            cells.task_cells,
            task_values.numbers,
            task_values.max_points,
            filled_rows.last_row_number,
        )

    def select_filled_rows(self, cell_range: Range) -> FilledRows:
        start_cell: Cell = self.ws[cell_range.start]
        end_cell: Cell = self.ws[cell_range.end]
        rows: List[Tuple[Cell, ...]] = []
        last_row_number = end_cell.row

        for row in self.ws.iter_rows(
            min_row=start_cell.row,
            max_row=end_cell.row,
            min_col=start_cell.column,
            max_col=end_cell.column,
        ):
            cell = self.ws.cell(row[0].row, row[0].column - 1)

            if cell.value != EMPTY_STUDENT:
                rows.append(row)

        return FilledRows(tuple(rows), last_row_number)

    def replace_x_cells(self, cell_range: MatrixCells) -> MatrixCells:
        for row in cell_range:
            for cell in row:
                if cell.value in REPLACE_VALUES:
                    cell.value = 0

        return cell_range

    def find_student_cells(self, point_cells: MatrixCells) -> LineCells:
        student_cells = [self.ws.cell(row=row[0].row, column=1) for row in point_cells]
        return tuple(student_cells)

    def find_task_cells(self, point_cells: MatrixCells) -> LineCells:
        tasks_row = 1
        start_column = point_cells[0][0].column
        end_column = point_cells[0][-1].column
        return tuple(
            self.ws.cell(tasks_row, col) for col in range(start_column, end_column + 1)
        )

    def cell_range_finder(self, cell_range: MatrixCells) -> FinderCells:
        student_cells = self.find_student_cells(cell_range)
        task_cells = self.find_task_cells(cell_range)
        return FinderCells(student_cells, task_cells)

    def select_task_values(self, cell_range: LineCells) -> TaskValues:
        numbers = []
        max_scores = []

        for cell in cell_range:
            if type(cell.value) is str:
                number, max_score = cell.value.split(" ")
                numbers.append(number)
                max_score = max_score[1:-2]
                max_scores.append(int(max_score))
            else:
                raise ValueError("Cell value is not a string")

        return TaskValues(tuple(numbers), tuple(max_scores))
