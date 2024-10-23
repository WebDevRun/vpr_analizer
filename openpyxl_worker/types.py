from dataclasses import dataclass
from enum import Enum
from typing import List, Literal, NamedTuple, Tuple

from openpyxl.cell.cell import Cell

LineCells = Tuple[Cell, ...]
MatrixCells = Tuple[LineCells, ...]


class NumberFormatCell(Enum):
    FORMAT_PERCENTAGE_00 = "0.00%"
    FORMAT_NUMBER_00 = "0.00"
    NONE = None


class TableHeader(NamedTuple):
    number: str
    verifiable_requirements: str
    max_point: str
    average_point: str
    percentage_of_completion: str


class ResultTableHeaders(NamedTuple):
    number: str
    class_name: str
    task_number: str
    task_name: str
    percentage_of_completion: str


@dataclass
class Range:
    start: str
    end: str


@dataclass
class WsData:
    name: str
    point_range: Range


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
    task_formulas: LineCells
    student_formulas: LineCells
    point_formulas: MatrixCells
    average_formulas: LineCells
    percentage_of_completion_formulas: LineCells
    max_point_cells: LineCells
    sum_max_point_formula: Cell
    sum_student_point_formulas: LineCells
    average_point: Cell
    average_percentage_of_completion: Cell
    percentage_of_points: LineCells


@dataclass
class OverallResult:
    number: Cell
    class_number: Cell
    task_number: Cell
    task_name: Cell
    percentage_of_completion: Cell


@dataclass
class ResultCells:
    name: str
    overall_result: List[OverallResult]
