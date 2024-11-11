from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook


class MiscUtility:

    @staticmethod
    def open_workbook(file_path: str) -> Workbook:
        workbook = load_workbook(filename=file_path)

        return workbook

    def close_workbook(workbook: Workbook):
        if workbook is not None:
            workbook.close()

    @staticmethod
    def format_array_as_bullets(items: list[str]) -> str:
        sep = "\n\t- "

        return sep + sep.join(items)
