from dataclasses import asdict
from pathlib import Path
from typing import List

from yaml import dump, safe_load

from openpyxl_worker.types import Range
from yaml_worker.types import Workbook, WorkbooksRanges, Worksheet


class YamlWorker:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.table_config_path = Path("config", "table_ranges.yaml")

    def read(self) -> List[Workbook]:
        with open(self.path, "r", encoding="utf-8") as file:
            yaml = safe_load(file)

        workbooks_yaml = yaml.get("workbooks")
        workbooks: List[Workbook] = []

        for wb in workbooks_yaml:
            worksheets_data = wb["worksheets"]
            worksheets: List[Worksheet] = []

            for ws in worksheets_data:
                point_range = ws["point_range"]
                [start, end] = point_range.split(":")

                worksheets.append(Worksheet(ws["name"], Range(start, end)))

            workbooks.append(Workbook(wb["name"], worksheets))

        return workbooks

    def write(self, workbooks_ranges: WorkbooksRanges) -> None:
        with open(self.table_config_path, "w", encoding="utf-8") as file:
            dump(asdict(workbooks_ranges), file)
