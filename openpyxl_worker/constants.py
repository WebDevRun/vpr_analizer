"""Constants for openpyxl_worker: styles, colors, and table headers for Excel processing.

This module defines reusable constants for table headers, cell alignments, colors, and borders used in Excel workbook processing.
"""

from openpyxl.styles.borders import BORDER_THIN, Border, Side

from openpyxl_worker.types import AlignmentCell, ResultTableHeaders, TableHeader

THEME_TABLE_HEADERS: TableHeader = TableHeader(
    "№",
    "Проверяемые требования",
    "Максимальный балл",
    "Средний балл",
    "Процент выполнения",
)
"""Default table headers for theme tables."""

THEME_RESULT_TABLE_HEADERS: ResultTableHeaders = ResultTableHeaders(
    "№",
    "Номер задания",
    "Требования",
    "Процент выполнения",
)
"""Default headers for result tables by theme."""

LEFT_TOP_ALIGN: AlignmentCell = AlignmentCell("left", "top")
"""Cell alignment: left horizontally, top vertically."""

RIGHT_TOP_ALIGN: AlignmentCell = AlignmentCell("right", "top")
"""Cell alignment: right horizontally, top vertically."""

LIME_COLOR: str = "81d41a"
"""Hex color code for lime highlight."""

YELLOW_COLOR: str = "ffff00"
"""Hex color code for yellow highlight."""

BRICK_COLOR: str = "ff4000"
"""Hex color code for brick highlight."""

THIN_BORDER: Border = Border(
    left=Side(style=BORDER_THIN),
    right=Side(style=BORDER_THIN),
    top=Side(style=BORDER_THIN),
    bottom=Side(style=BORDER_THIN),
)
"""Thin border style for table cells."""

SUMMARY_TABLE_TITLE: str = "Общие_результаты"
"""Default title for summary tables."""
