from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
import openpyxl


class MiscUtility:

    @staticmethod
    def open_workbook(file_path: str) -> Workbook:
        workbook = load_workbook(filename=file_path)
        # workbook.iso_dates = True
        # workbook.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904

        return workbook

    def close_workbook(workbook: Workbook):
        if workbook is not None:
            workbook.close()

    @staticmethod
    def format_array_as_bullets(items: list[str]) -> str:
        if len(items) == 0:
            return ""

        sep = "\n\t- "

        return sep + sep.join(items)
