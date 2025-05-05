import logging
import os
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


def main() -> None:
    """Main entry point for the VPR analyzer application.

    Reads workbook configurations, processes each worksheet, creates analytic and summary tables, and saves results.
    Enhanced with error handling and logging.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    table_config_path = Path(os.getenv("TABLES_CONFIG_PATH", "tables.yaml"))
    try:
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
                logging.info("%s %s - %s", Sentences.create_table, wb.name, ws.name)

            SummaryTableWorker(wb_container.wb, SUMMARY_TABLE_TITLE).create(
                summary_table_data
            )
            logging.info(
                "%s %s - %s", Sentences.create_table, wb.name, Sentences.overall_results
            )
            wb_container.save_table(table_path)
            logging.info("%s %s", Sentences.save_table, table_path)

        # For CLI use, uncomment the next line:
        # input(Sentences.press_to_close)
    except Exception as exc:
        logging.exception("An error occurred during processing: %s", exc)


if __name__ == "__main__":
    main()
