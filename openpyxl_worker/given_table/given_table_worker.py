import logging

from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_worker.given_table.cell_utils import (
    extract_student_and_task_cells,
    get_nonempty_rows,
    remove_variant_columns,
    replace_cells_with_zero,
)
from openpyxl_worker.types import (
    GivenTableCells,
    LineCells,
    Range,
    TaskValues,
)


class GivenTableWorker:
    """
    Worker class for extracting and processing given table data from an Excel worksheet.
    Handles filled row selection, cell value replacement, variant column skipping, and cell range finding.
    """

    OPEN_SCORE = "("
    VARIANT = "вариант"

    def __init__(self, ws: Worksheet, point_range: Range) -> None:
        """Initialize the GivenTableWorker.

        Args:
            ws (Worksheet): The worksheet to process.
            point_range (Range): The cell range to process.
        """
        self.ws = ws
        self.point_range = point_range

    def get_cell_ranges(self) -> GivenTableCells:
        """Extract and process all relevant cell ranges for the given table.

        Returns:
            GivenTableCells: Named tuple containing all relevant cell ranges and values.
        """
        filled_rows = get_nonempty_rows(self.ws, self.point_range)
        point_cells = replace_cells_with_zero(filled_rows.rows)
        point_cells = remove_variant_columns(self.ws, point_cells, self.VARIANT)
        cells = extract_student_and_task_cells(self.ws, filled_rows.rows)
        task_values = self.select_task_values(cells.task_cells)
        logging.info("Extracted cell ranges for given table.")
        return GivenTableCells(
            point_cells,
            cells.student_cells,
            cells.task_cells,
            task_values.numbers,
            task_values.max_points,
            filled_rows.last_row_number,
        )

    def select_task_values(self, cell_range: LineCells) -> TaskValues:
        """Extract task numbers and max scores from the task cell range.

        Args:
            cell_range (LineCells): The line of task cells to process.

        Returns:
            TaskValues: Named tuple of task numbers and max scores.
        """
        numbers = []
        max_scores = []

        for cell in cell_range:
            if isinstance(cell.value, str):
                if self.VARIANT in cell.value.lower():
                    continue
                strip_str = cell.value.strip()
                try:
                    score_index = strip_str.index(self.OPEN_SCORE)
                    number = strip_str[:score_index]
                    max_score = strip_str[score_index + 1 : -2]
                    numbers.append(number)
                    max_scores.append(int(max_score))
                except (ValueError, IndexError):
                    logging.error(
                        "Failed to parse task value from cell '%s': %s",
                        cell.coordinate,
                        cell.value,
                    )
                    continue

        logging.info("Selected %d task values.", len(numbers))
        return TaskValues(tuple(numbers), tuple(max_scores))
