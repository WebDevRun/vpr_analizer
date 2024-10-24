from openpyxl_worker.types import ResultTableHeaders, TableHeader

THEME_TABLE_HEADERS = TableHeader(
    "№",
    "Проверяемые требования",
    "Максимальный балл",
    "Средний балл",
    "Процент выполнения",
)

REPLACE_VALUES = ("x", "X", "х", "Х")

THEME_RESULT_TABLE_HEADERS = ResultTableHeaders(
    "№",
    "Класс",
    "Номер задания",
    "Требования",
    "Процент выполнения",
)
