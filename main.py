from pathlib import Path
from typing import List

from openpyxl_worker import (
    AnalyticTableCreates,
    GivenTableWorker,
    SummaryTableWorker,
    WorkbookContainer,
    WorksheetRanges,
)
from openpyxl_worker.constants import SUMMARY_TABLE_TITLE
from sentences import Directory, Sentences
from yaml_worker import YamlWorker


def main():
    table_config_path = Path("tables.yaml")
    yaml_worker = YamlWorker(table_config_path)
    workbooks = yaml_worker.read()

    for wb in workbooks:
        table_path = Path(Directory.tables, wb.name)
        wb_container = WorkbookContainer(table_path)
        summary_table_data: List[WorksheetRanges] = []

        for ws in wb.worksheets:
            wb_data = wb_container.activate_sheet(ws.name)
            given_ranges = GivenTableWorker(
                wb_data.ws, ws.point_range
            ).get_cell_ranges()
            worksheet_ranges = AnalyticTableCreates(
                wb_data.wb, wb_data.ws, given_ranges
            ).create()
            summary_table_data.append(worksheet_ranges)
            print(f"{Sentences.create_table} {wb.name} - {ws.name}")

        SummaryTableWorker(wb_container.wb, SUMMARY_TABLE_TITLE).create(
            summary_table_data
        )
        print(f"{Sentences.create_table} {wb.name} - {Sentences.overall_results}")
        wb_container.save_table(table_path)
        print(f"{Sentences.save_table} {table_path}")

    input(Sentences.press_to_close)


if __name__ == "__main__":
    main()
