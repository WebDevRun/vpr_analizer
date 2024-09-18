from typing import List, Tuple

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_worker.cells_finder import MatrixCells
from openpyxl_worker.constants import THEME_TABLE_HEADERS
from openpyxl_worker.types import LineCells


class CellsFiller:
    def __init__(self, ws: Worksheet):
        self.ws = ws

    def fill_table_header(self, student_cells: LineCells) -> LineCells:
        cell = self.ws.cell(student_cells[-1].row + 2, 1, THEME_TABLE_HEADERS.number)
        filled_cells: List[Cell] = [cell]
        student_formulas = tuple([f"={cell.coordinate}" for cell in student_cells])
        headers = (
            THEME_TABLE_HEADERS[1:3]
            + student_formulas
            + THEME_TABLE_HEADERS[3 : len(THEME_TABLE_HEADERS)]
        )

        for header in headers:
            cell = self.ws.cell(cell.row, cell.column + 1, header)
            filled_cells.append(cell)

        return tuple(filled_cells)

    def fill_task_numbers(self, task_cells: LineCells, number_cell: Cell) -> LineCells:
        start_cell = self.ws.cell(number_cell.row + 1, number_cell.column)
        filled_cells: List[Cell] = []

        for index, task_cell in enumerate(task_cells):
            cell_formula = f"={task_cell.coordinate}"
            cell = self.ws.cell(start_cell.row + index, start_cell.column, cell_formula)
            filled_cells.append(cell)
        return tuple(filled_cells)

    def fill_point_formulas(
        self, point_cells: MatrixCells, student_cells: LineCells
    ) -> MatrixCells:
        filled_cells: List[Tuple[Cell, ...]] = []

        for index, point_cell_row in enumerate(point_cells):
            column = student_cells[index].column
            start_row = student_cells[index].row + 1
            filled_row = []

            for index, point_cell in enumerate(point_cell_row):
                cell_formula = f"={point_cell.coordinate}"
                cell = self.ws.cell(start_row + index, column, cell_formula)
                filled_row.append(cell)

            filled_cells.append(tuple(filled_row))

            column += 1

        return tuple(filled_cells)

    def fill_average_formulas(
        self, point_cells: MatrixCells, average_cell: Cell
    ) -> LineCells:
        column = average_cell.column
        start_row = average_cell.row + 1
        filled_cells: List[Cell] = []

        for index, cell in enumerate(point_cells[0]):
            start_cell_coordinate = f"{cell.column_letter }{cell.row}"
            last_cell = point_cells[-1][index]
            end_cell_coordinate = f"{last_cell.column_letter}{last_cell.row}"
            cell_formula = f"=AVERAGE({start_cell_coordinate}:{end_cell_coordinate})"
            cell = self.ws.cell(start_row + index, column, cell_formula)
            filled_cells.append(cell)

        return tuple(filled_cells)

    def fill_percentage_of_completion(
        self,
        average_cells: LineCells,
        max_point_cell: Cell,
        percentage_of_completion_cell: Cell,
    ) -> LineCells:
        column = percentage_of_completion_cell.column
        start_row = percentage_of_completion_cell.row + 1
        filled_cells: List[Cell] = []

        for index, cell in enumerate(average_cells):
            max_point_coordinate = f"{max_point_cell.column_letter}{cell.row}"
            cell_formula = f"={cell.coordinate}/{max_point_coordinate}"
            cell = self.ws.cell(start_row + index, column, cell_formula)
            filled_cells.append(cell)

        return tuple(filled_cells)

    def fill_sum_max_points(self, max_point_cells: LineCells):
        last_cell = max_point_cells[-1]
        row = last_cell.row + 1
        column = last_cell.column
        sum_max_point_formula = self.ws.cell(row, column)
        sum_max_point_formula.value = (
            f"=SUM({max_point_cells[0].coordinate}:{max_point_cells[-1].coordinate}"
        )
        return sum_max_point_formula

    def fill_sum_student_points(self, student_cells: MatrixCells):
        filled_cells: List[Cell] = []

        for student_cell_row in student_cells:
            last_cell = student_cell_row[-1]
            row = last_cell.row + 1
            column = last_cell.column
            sum_student_point_formula = self.ws.cell(row, column)
            sum_student_point_formula.value = f"=SUM({student_cell_row[0].coordinate}:{student_cell_row[-1].coordinate})"
            filled_cells.append(sum_student_point_formula)

        return tuple(filled_cells)

    def fill_average_point(self, average_formulas: Tuple[Cell, ...]):
        last_cell = average_formulas[-1]
        row = last_cell.row + 1
        column = last_cell.column
        cell = self.ws.cell(row, column)
        cell.value = f"=AVERAGE({average_formulas[0].coordinate}:{average_formulas[-1].coordinate})"
        return cell

    def fill_average_percentage_of_completion(
        self,
        sum_student_point_formulas: Tuple[Cell, ...],
        sum_max_point_formula: Cell,
        last_percentage_of_completion_formula: Cell,
    ):
        row = last_percentage_of_completion_formula.row + 1
        column = last_percentage_of_completion_formula.column
        cell = self.ws.cell(row, column)
        average_chunk = f"=AVERAGE({sum_student_point_formulas[0].coordinate}:{sum_student_point_formulas[-1].coordinate})"
        cell.value = f"{average_chunk}/{sum_max_point_formula.coordinate}"
        return cell

    def fill_percentage_of_point(
        self, sum_student_point_formula: Tuple[Cell, ...], sum_max_point_formula: Cell
    ):
        filled_cell: List[Cell] = []

        for cell in sum_student_point_formula:
            percent_cell = self.ws.cell(cell.row + 1, cell.column)
            percent_cell.value = (
                f"={cell.coordinate}/{sum_max_point_formula.coordinate}"
            )
            filled_cell.append(percent_cell)

        return tuple(filled_cell)
