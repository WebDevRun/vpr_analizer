from enum import StrEnum
from pathlib import Path
from typing import List

from yaml import safe_load

from argparser import ArgParser
from openpyxl_worker import RangeConverter, TableWorker
from yaml_worker import YamlWorker
from yaml_worker.types import WorkbookRanges, WorkbooksRanges, WorksheetStrRanges


class Directory(StrEnum):
    config = "config"
    tables = "tables"


class Sentences(StrEnum):
    conditional_formatting = "Создано условное форматирование для таблицы:"
    create_table = "Создана таблица:"
    save_table = "Сохранена таблица:"
    press_to_close = "Нажмите Enter для закрытия..."


def main():
    args = ArgParser().get_args()

    if args.paint:
        table_config_path = Path(Directory.config.value, "table_ranges.yaml")
        with open(table_config_path, "r", encoding="utf-8") as file:
            yaml = safe_load(file)

        for wb in yaml.get("workbooks"):
            table_name = wb.get("name")
            table_path = Path(Directory.tables.value, table_name)
            table_worker = TableWorker(table_path)

            for worksheet in wb.get("worksheets"):
                ws_name = worksheet.get("name")
                ws = table_worker.activate_sheet(ws_name)
                ranges = RangeConverter(table_worker.ws).str_to_ranges(worksheet)
                table_worker.paint_worksheet(ranges)
                print(
                    f"{Sentences.conditional_formatting.value} {table_name} - {ws_name}"
                )

            table_worker.save_table(table_path)
            print(f"{Sentences.save_table.value} {table_path}")

        input(Sentences.press_to_close.value)
        return

    table_config_path = Path("tables.yaml")

    yaml_worker = YamlWorker(table_config_path)
    workbooks = yaml_worker.read()
    workbook_list: List[WorkbookRanges] = []

    for wb in workbooks:
        table_path = Path(Directory.tables.value, wb.name)
        table_worker = TableWorker(table_path)
        worksheet_list: List[WorksheetStrRanges] = []

        for ws in wb.worksheets:
            worksheet_ranges = (
                table_worker.activate_sheet(ws.name)
                .replace_x_cells(ws.point_range)
                .fill_theme_table(ws.point_range)
            )
            worksheet_str_ranges = RangeConverter(table_worker.ws).ranges_to_str(
                worksheet_ranges
            )
            worksheet_list.append(worksheet_str_ranges)
            print(f"{Sentences.create_table.value} {wb.name} - {ws.name}")

        workbook_list.append(WorkbookRanges(wb.name, worksheet_list))
        table_worker.save_table(table_path)
        print(f"{Sentences.save_table.value} {table_path}")

    yaml_worker.write(WorkbooksRanges(workbook_list))
    input(Sentences.press_to_close.value)


if __name__ == "__main__":
    main()
