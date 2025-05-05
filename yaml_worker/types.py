from dataclasses import dataclass
from typing import List

from openpyxl_worker.types import Range

Name = str


@dataclass
class Worksheet:
    """Represents a worksheet configuration with a name and a point range.

    Attributes:
        name (Name): The name of the worksheet.
        point_range (Range): The cell range for the worksheet (e.g., C2:R21).
    """

    name: Name
    point_range: Range


@dataclass
class Workbook:
    """Represents a workbook configuration with a name and a list of worksheets.

    Attributes:
        name (Name): The name of the workbook (Excel file).
        worksheets (List[Worksheet]): List of worksheet configurations in the workbook.
    """

    name: Name
    worksheets: List[Worksheet]


@dataclass
class WorksheetStrRanges:
    """Represents string-based cell ranges and formulas for a worksheet.

    Attributes:
        name (str): Worksheet name.
        table_headers (str): Table header cell references.
        task_formulas (str): Task formula cell references.
        student_formulas (str): Student formula cell references.
        point_formulas (str): Point formula cell references.
        average_formulas (str): Average formula cell references.
        percentage_of_completion_formulas (str): Completion percentage formula cell references.
        max_point_cells (str): Max point cell references.
        sum_max_point_formula (str): Formula for sum of max points.
        sum_student_point_formulas (str): Formula for sum of student points.
        average_point (str): Average point cell reference.
        average_percentage_of_completion (str): Average completion percentage cell reference.
        percentage_of_points (str): Percentage of points cell reference.
    """

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
    """Represents a workbook with string-based worksheet ranges.

    Attributes:
        name (Name): The name of the workbook.
        worksheets (List[WorksheetStrRanges]): List of worksheet string range definitions.
    """

    name: Name
    worksheets: List[WorksheetStrRanges]


@dataclass
class WorkbooksRanges:
    """Represents a collection of workbook ranges.

    Attributes:
        workbooks (List[WorkbookRanges]): List of workbook range definitions.
    """

    workbooks: List[WorkbookRanges]
