from dataclasses import dataclass
from enum import Enum
from typing import List, Literal, NamedTuple, Tuple

from openpyxl.cell.cell import Cell

LineCells = Tuple[Cell, ...]
MatrixCells = Tuple[LineCells, ...]


class TableHeader(NamedTuple):
    number: str
    verifiable_requirements: str
    max_point: str
    average_point: str
    percentage_of_completion: str


class ResultTableHeaders(NamedTuple):
    number: str
    class_name: str
    task_name: str
    percentage_of_completion: str


@dataclass
class Range:
    start: str
    end: str


class NumberFormatCell(Enum):
    FORMAT_PERCENTAGE_00 = "0.00%"
    FORMAT_NUMBER_00 = "0.00"
    NONE = None


@dataclass
class AlignmentCell:
    horizontal: Literal["right", "center", "left"]
    vertical: Literal["top", "center", "bottom"]


@dataclass
class FormatArgs:
    alignment: AlignmentCell
    number_format: NumberFormatCell = NumberFormatCell.NONE
    wrap_text: bool = False


@dataclass
class WorksheetRanges:
    name: str
    table_headers: LineCells
    task_cells: LineCells
    point_formulas: MatrixCells
    average_formulas: LineCells
    percentage_of_completion_formulas: LineCells
    max_point_cells: LineCells
    average_point: Cell
    average_percentage_of_completion: Cell
    percentage_of_points: LineCells
    task_discription_cells: LineCells


@dataclass
class OverallResult:
    number: Cell
    task_number: Cell
    task_name: Cell
    percentage_of_completion: Cell


@dataclass
class ResultCells:
    name: str
    overall_result: List[OverallResult]


@dataclass
class GivenTableCells:
    point_cells: MatrixCells
    student_cells: LineCells
    task_cells: LineCells
    task_numbers: Tuple[str, ...]
    max_points: Tuple[int, ...]
    last_row: int


@dataclass
class FinderCells:
    student_cells: LineCells
    task_cells: LineCells


@dataclass
class FilledRows:
    rows: MatrixCells
    last_row_number: int


@dataclass
class TaskValues:
    numbers: Tuple[str, ...]
    max_points: Tuple[int, ...]
