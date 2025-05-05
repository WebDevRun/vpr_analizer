from dataclasses import dataclass
from enum import Enum
from typing import List, Literal, NamedTuple, Tuple

from openpyxl.cell.cell import Cell

LineCells = Tuple[Cell, ...]
MatrixCells = Tuple[LineCells, ...]


class TableHeader(NamedTuple):
    """Named tuple for table header fields."""

    number: str
    verifiable_requirements: str
    max_point: str
    average_point: str
    percentage_of_completion: str


class ResultTableHeaders(NamedTuple):
    """Named tuple for result table header fields."""

    number: str
    class_name: str
    task_name: str
    percentage_of_completion: str


@dataclass
class Range:
    """Represents a cell range with start and end cell references."""

    start: str
    end: str


class NumberFormatCell(Enum):
    FORMAT_PERCENTAGE_00 = "0.00%"
    FORMAT_NUMBER_00 = "0.00"
    NONE = None


@dataclass
class AlignmentCell:
    """Represents cell alignment options for Excel cells."""

    horizontal: Literal["right", "center", "left"]
    vertical: Literal["top", "center", "bottom"]


@dataclass
class FormatArgs:
    """Arguments for formatting Excel cells."""

    alignment: AlignmentCell
    number_format: NumberFormatCell = NumberFormatCell.NONE
    wrap_text: bool = False


@dataclass
class WorksheetRanges:
    """Represents all relevant cell ranges for a worksheet."""

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
    """Represents a row in the summary table with all relevant cells."""

    number: Cell
    task_number: Cell
    task_name: Cell
    percentage_of_completion: Cell


@dataclass
class ResultCells:
    """Represents all result rows for a summary table."""

    name: str
    overall_result: List[OverallResult]


@dataclass
class GivenTableCells:
    """Represents all relevant cell ranges and values for a given table."""

    point_cells: MatrixCells
    student_cells: LineCells
    task_cells: LineCells
    task_numbers: Tuple[str, ...]
    max_points: Tuple[int, ...]
    last_row: int


@dataclass
class FinderCells:
    """Represents found student and task cells for a matrix."""

    student_cells: LineCells
    task_cells: LineCells


@dataclass
class FilledRows:
    """Represents filled rows and the last row number in a worksheet."""

    rows: MatrixCells
    last_row_number: int


@dataclass
class TaskValues:
    """Represents extracted task numbers and max points from task cells."""

    numbers: Tuple[str, ...]
    max_points: Tuple[int, ...]
