from openpyxl.styles import Font


class ExcelCellData:

    def __init__(self, value, font: Font = None):
        self.value = value
        self.font = font
