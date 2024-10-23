from enum import StrEnum
from pathlib import Path
from typing import List

from yaml import safe_load

from argparser import ArgParser
from openpyxl_worker import RangeConverter, TableWorker
from openpyxl_worker.overall_results_worker import OverallResultsWorker
from openpyxl_worker.types import WorksheetRanges
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
    overall_results = "Общие результаты"


def main():
    args = ArgParser().get_args()

    if args.paint:
        table_config_path = Path(Directory.config, "table_ranges.yaml")
        with open(table_config_path, "r", encoding="utf-8") as file:
            yaml = safe_load(file)

        for wb in yaml.get("workbooks"):
            table_name = wb.get("name")
            table_path = Path(Directory.tables, table_name)
            table_worker = TableWorker(table_path)

            for worksheet in wb.get("worksheets"):
                ws_name = worksheet.get("name")
                ws = table_worker.activate_sheet(ws_name)
                ranges = RangeConverter(table_worker.ws).str_to_ranges(worksheet)
                table_worker.paint_worksheet(ranges)
                print(
                    f"{Sentences.conditional_formatting} {table_name} - {ws_name}"
                )

            table_worker.save_table(table_path)
            print(f"{Sentences.save_table} {table_path}")

        input(Sentences.press_to_close)
        return

    table_config_path = Path("tables.yaml")

    yaml_worker = YamlWorker(table_config_path)
    workbooks = yaml_worker.read()
    workbook_list: List[WorkbookRanges] = []

    for wb in workbooks:
        table_path = Path(Directory.tables, wb.name)
        table_worker = TableWorker(table_path)
        wb_ranges: List[WorksheetRanges] = []
        worksheet_list: List[WorksheetStrRanges] = []

        for ws in wb.worksheets:
            worksheet_ranges = (
                table_worker.activate_sheet(ws.name)
                .replace_x_cells(ws.point_range)
                .fill_theme_table(ws.point_range)
            )
            wb_ranges.append(worksheet_ranges)
            worksheet_str_ranges = RangeConverter(table_worker.ws).ranges_to_str(
                worksheet_ranges
            )
            worksheet_list.append(worksheet_str_ranges)
            print(f"{Sentences.create_table} {wb.name} - {ws.name}")

        workbook_list.append(WorkbookRanges(wb.name, worksheet_list))

        OverallResultsWorker(
            table_worker.wb, Sentences.overall_results
        ).fill_table(wb_ranges)
        print(
            f"{Sentences.create_table} {wb.name} - {Sentences.overall_results}"
        )

        table_worker.save_table(table_path)
        print(f"{Sentences.save_table} {table_path}")

    yaml_worker.write(WorkbooksRanges(workbook_list))
    input(Sentences.press_to_close)


if __name__ == "__main__":
    main()
