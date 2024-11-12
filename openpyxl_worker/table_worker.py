from pathlib import Path

from openpyxl import load_workbook


class WorkbookContainer:
    def __init__(self, file_path: Path):
        self.path = file_path
        self.wb = load_workbook(self.path)
        self.ws = self.wb[self.wb.sheetnames[0]]

    def activate_sheet(self, sheet_name: str):
        self.ws = self.wb[sheet_name]
        return self

    def save_table(self, file_path: Path):
        self.wb.save(file_path)
