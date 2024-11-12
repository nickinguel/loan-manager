import datetime
import os
from copy import copy

from dateutil import relativedelta
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
import openpyxl


class MiscUtility:

    @staticmethod
    def get_next_month(date: datetime.date) -> datetime.date:
        return date + relativedelta.relativedelta(months=1)
        next_month_date: datetime.datetime = copy(date)

        try:
            next_month_date = next_month_date.replace(month=next_month_date.month + 1)
        except ValueError:
            if next_month_date.month == 12:
                next_month_date = next_month_date.replace(year=next_month_date.year + 1, month=1)

        return next_month_date

    @staticmethod
    def open_workbook(file_path: str) -> Workbook:
        if not os.path.exists(file_path):
            raise Exception(f"Aucun fichier trouvé à l'emplacement spécifié '{file_path}'")

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
