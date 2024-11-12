from enum import StrEnum


class Directory(StrEnum):
    config = "config"
    tables = "tables"


class Sentences(StrEnum):
    conditional_formatting = "Создано условное форматирование для таблицы:"
    create_table = "Создана таблица:"
    save_table = "Сохранена таблица:"
    press_to_close = "Нажмите Enter для закрытия..."
    overall_results = "Общие результаты"
