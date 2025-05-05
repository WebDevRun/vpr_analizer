import logging
import os
from dataclasses import asdict
from pathlib import Path
from typing import List

from yaml import YAMLError, dump, safe_load

from openpyxl_worker.types import Range
from yaml_worker.types import Workbook, WorkbooksRanges, Worksheet


class YamlWorker:
    """Handles reading and writing workbook configurations in YAML format for Excel processing projects.

    This class provides methods to read workbook/worksheet structures from YAML and write processed ranges back to YAML.
    The output path for writing can be configured via the TABLE_CONFIG_OUTPUT_PATH environment variable.
    """

    def __init__(self, path: Path) -> None:
        """Initialize YamlWorker with the path to the YAML configuration file.

        Args:
            path (Path): Path to the YAML file containing workbook configurations.
        """
        self.path: Path = path
        self.table_config_path: Path = Path(
            os.getenv("TABLE_CONFIG_OUTPUT_PATH", "config/table_ranges.yaml")
        )

    def read(self) -> List[Workbook]:
        """Read workbook configurations from the YAML file.

        Returns:
            List[Workbook]: List of Workbook objects parsed from YAML.
        Raises:
            FileNotFoundError: If the YAML file does not exist.
            YAMLError: If the YAML file is invalid.
            OSError: For other I/O errors.
            ValueError: If worksheet point_range is malformed.
        """
        try:
            with open(self.path, "r", encoding="utf-8") as file:
                yaml_data = safe_load(file)
        except (FileNotFoundError, YAMLError, OSError):
            logging.exception("Failed to read or parse YAML file: %s", self.path)
            raise

        workbooks_yaml = yaml_data.get("workbooks", [])
        workbooks: List[Workbook] = []

        for wb in workbooks_yaml:
            worksheets_data = wb["worksheets"]
            worksheets: List[Worksheet] = []

            for ws in worksheets_data:
                point_range = ws["point_range"]
                try:
                    start, end = point_range.split(":")
                except ValueError as ve:
                    logging.error(
                        "Invalid point_range format in worksheet '%s' of workbook '%s': %s",
                        ws.get("name", "<unknown>"),
                        wb.get("name", "<unknown>"),
                        point_range,
                    )
                    raise ValueError(
                        f"Invalid point_range format: '{point_range}' in worksheet '{ws.get('name', '<unknown>')}' of workbook '{wb.get('name', '<unknown>')}'"
                    ) from ve
                worksheets.append(Worksheet(ws["name"], Range(start, end)))

            workbooks.append(Workbook(wb["name"], worksheets))

        logging.info(
            "Successfully read %d workbooks from %s", len(workbooks), self.path
        )
        return workbooks

    def write(self, workbooks_ranges: WorkbooksRanges) -> None:
        """Write workbook ranges to the YAML configuration file.

        Args:
            workbooks_ranges (WorkbooksRanges): The workbook ranges to serialize and write.
        Raises:
            OSError: If writing to the file fails.
        """
        try:
            with open(self.table_config_path, "w", encoding="utf-8") as file:
                dump(asdict(workbooks_ranges), file)
            logging.info(
                "Successfully wrote workbook ranges to %s", self.table_config_path
            )
        except OSError:
            logging.exception(
                "Failed to write workbook ranges to %s", self.table_config_path
            )
            raise
