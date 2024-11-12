from openpyxl.styles.borders import BORDER_THIN, Border, Side

from openpyxl_worker.types import AlignmentCell, ResultTableHeaders, TableHeader

THEME_TABLE_HEADERS = TableHeader(
    "№",
    "Проверяемые требования",
    "Максимальный балл",
    "Средний балл",
    "Процент выполнения",
)

THEME_RESULT_TABLE_HEADERS = ResultTableHeaders(
    "№",
    "Номер задания",
    "Требования",
    "Процент выполнения",
)

LEFT_TOP_ALIGN = AlignmentCell("left", "top")
RIGHT_TOP_ALIGN = AlignmentCell("right", "top")

LIME_COLOR = "81d41a"
YELLOW_COLOR = "ffff00"
BRICK_COLOR = "ff4000"

THIN_BORDER = Border(
    left=Side(style=BORDER_THIN),
    right=Side(style=BORDER_THIN),
    top=Side(style=BORDER_THIN),
    bottom=Side(style=BORDER_THIN),
)

SUMMARY_TABLE_TITLE = "Общие_результаты"
