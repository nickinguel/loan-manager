import datetime
import os
from dateutil import relativedelta
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename


class MiscUtility:

    @staticmethod
    def browse_file_path(prompt: str):
        Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        filename = askopenfilename(
            title=prompt,
            filetypes=(('Excel', "*.xlsx"), ('Excel legacy', "*.xls"))
        )

        return filename

    @staticmethod
    def get_next_month(date: datetime.date) -> datetime.date:
        return date + relativedelta.relativedelta(months=1)

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

        sep = "\n  - "

        return sep + sep.join(items)
