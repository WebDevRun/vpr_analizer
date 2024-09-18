from dataclasses import dataclass
from typing import List

from openpyxl_worker.types import Range

Name = str


@dataclass
class Worksheet:
    name: Name
    point_range: Range


@dataclass
class Workbook:
    name: Name
    worksheets: List[Worksheet]


@dataclass
class WorksheetStrRanges:
    name: str
    table_headers: str
    task_formulas: str
    student_formulas: str
    point_formulas: str
    average_formulas: str
    percentage_of_completion_formulas: str
    max_point_cells: str
    sum_max_point_formula: str
    sum_student_point_formulas: str
    average_point: str
    average_percentage_of_completion: str
    percentage_of_points: str


@dataclass
class WorkbookRanges:
    name: Name
    worksheets: List[WorksheetStrRanges]


@dataclass
class WorkbooksRanges:
    workbooks: List[WorkbookRanges]
