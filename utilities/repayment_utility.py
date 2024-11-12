from typing import Any

from openpyxl.cell import Cell
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.dimensions import DimensionHolder, ColumnDimension
from openpyxl.worksheet.worksheet import Worksheet
from configs.constant_data import ConstData
from models.excel_cell_data import ExcelCellData
from models.loan import Loan
from models.repayment import Repayment
from utilities.excel_utility import ExcelUtility
from utilities.misc_utility import MiscUtility


class RepaymentUtility:
    centered_alignment = Alignment(horizontal='center', vertical='center')

    @staticmethod
    def write_repayments_to_excel(file_path: str, data: dict[int, dict[str, ExcelCellData]]):
        """
        Write to file
        :param file_path:
        :param data:
        :return:
        """
        workbook = MiscUtility.open_workbook(file_path)
        sheet_prefix = ConstData.excel_sheet_repayment

        try:
            for year, cells_data in data.items():
                sheet_name = f"{sheet_prefix} {year}"
                sheet: Worksheet = workbook.get_sheet_by_name(sheet_name) if sheet_name in workbook.sheetnames else workbook.create_sheet(sheet_name)

                for cell_index, cell_data in cells_data.items():
                    cell: Cell = sheet[cell_index]

                    cell.alignment = RepaymentUtility.centered_alignment
                    cell.value = cell_data.value

                    if cell_data.font:
                        cell.font = cell_data.font

                RepaymentUtility.stylize_sheet(sheet)

            workbook.save(file_path)
        except (Exception,):
            raise
        finally:
            MiscUtility.close_workbook(workbook)

    @staticmethod
    def stylize_sheet(sheet: Worksheet):
        width: int

        # FONTS
        headers_row: DimensionHolder = sheet.row_dimensions[1]

        # SIZES
        # - Widths
        for index, column_name in enumerate(ConstData.excel_cols_repayments):
            if index == 0:
                width = 15
            elif index < 3:   # Loanee details
                width = 20
            elif index == 3:    # Amount
                width = 30
            else:               # Monts
                width = 20

            column_index_letter = ConstData.alphabet[index]

            column_dimension_holder: ColumnDimension = sheet.column_dimensions[column_index_letter]
            column_dimension_holder.width = width

            # Headers
            header_cell: Cell = sheet[f"{column_index_letter}1"]

            header_cell.alignment = RepaymentUtility.centered_alignment
            header_cell.font = Font(
                size=12,
                bold=True
            )

        # -- Heights
        headers_row.height = 40

        for i in range(2, sheet.max_row + 1):
            sheet.row_dimensions[i].height = 30

    @staticmethod
    def group_repayments_by_year(repayments: list[Repayment]):
        """
        Simply split and group repayments per year into a dictionary having as key te year and as value repayments
        :param repayments:
        :return:
        """
        repayments_grouped: dict[int, list[Repayment]] = {}

        for index, repayment in enumerate(repayments):
            years = set([sl.date.year for sl in repayment.slices])

            for year in years:
                if repayments_grouped.get(year) is None:
                    repayments_grouped[year] = []

                new_repayment = Repayment(repayment.loanee)
                new_repayment.slices = [sl for sl in repayment.slices if sl.date.year == year]

                repayments_grouped[year].append(new_repayment)

        repayments_grouped = dict(sorted(repayments_grouped.items()))

        return repayments_grouped

    @staticmethod
    def compute_sheets_data(repayments_grouped: dict[int, list[Repayment]]):
        """
        After computing loans to repayment object, this method compute repayments to sheets data into a dictionary
        :param repayments_grouped:
        :return:
        """

        sheets_data: dict[int, dict[str, ExcelCellData]] = {}

        for year, repayments in repayments_grouped.items():
            index = 0

            for repayment in repayments:
                index += 1

                for the_slice in repayment.slices:

                    if sheets_data.get(year) is None:
                        sheets_data[year] = {}
                        RepaymentUtility.fill_headers_cells(sheets_data, year)

                    RepaymentUtility.fill_loanee_cells(sheets_data, year, repayment, index)

                    month = ConstData.months[the_slice.date.month - 1]

                    cell_index = ExcelUtility.get_repayment_cell_from_column_name(month, index)
                    sheets_data[year][cell_index] = ExcelCellData(the_slice.amount)

        sheets_data = dict(sorted(sheets_data.items()))

        return sheets_data

    # @staticmethod
    # def compute_sheets_data(repayments: list[Repayment]):
    #     """
    #     After computing loans to repayment object, this method compute repayments to sheets data into a dictionary
    #     :param repayments:
    #     :return:
    #     """
    #
    #     sheets_data: dict[int, dict[str, Any]] = {}
    #
    #     for index, repayment in enumerate(repayments):
    #         year = -1
    #
    #         for the_slice in repayment.slices:
    #             if year != the_slice.date.year:
    #                 year = the_slice.date.year
    #
    #                 if sheets_data.get(year) is None:
    #                     sheets_data[year] = {}
    #                     RepaymentUtility.fill_headers_cells(sheets_data, year)
    #
    #                 RepaymentUtility.fill_loanee_cells(sheets_data, year, repayment, index + 1)
    #
    #             month = ConstData.months[the_slice.date.month - 1]
    #
    #             cell_index = ExcelUtility.get_repayment_cell_from_column_name(month, index + 1)
    #             sheets_data[year][cell_index] = the_slice.amount
    #
    #     sheets_data = dict(sorted(sheets_data.items()))
    #
    #     return sheets_data

    @staticmethod
    def fill_headers_cells(sheets_data: dict[int, dict[str, ExcelCellData]], year: int):
        for index, column_name in enumerate(ConstData.excel_cols_repayments):
            sheets_data[year][f"{ConstData.alphabet[index]}1"] = ExcelCellData(column_name)

    @staticmethod
    def fill_loanee_cells(sheets_data: dict[int, dict[str, ExcelCellData]], year: int, repayment: Repayment, index: int):

        # ID
        cell_index = ExcelUtility.get_repayment_cell_from_column_name(ConstData.excel_col_loan_debtor_ID, index)
        sheets_data[year][cell_index] = ExcelCellData(repayment.loanee.ID)

        # Firstname
        cell_index = ExcelUtility.get_repayment_cell_from_column_name(ConstData.excel_col_loan_debtor_first_name, index)
        sheets_data[year][cell_index] = ExcelCellData(repayment.loanee.first_name)

        # Lastname
        cell_index = ExcelUtility.get_repayment_cell_from_column_name(ConstData.excel_col_loan_debtor_last_name, index)
        sheets_data[year][cell_index] = ExcelCellData(repayment.loanee.last_name)

        # Total
        cell_index = ExcelUtility.get_repayment_cell_from_column_name(ConstData.excel_col_repayment_total_amount_loaned, index)
        sheets_data[year][cell_index] = ExcelCellData(sum([slice.amount for slice in repayment.slices]), Font(color="104db0", bold=False))

    @staticmethod
    def compute_repayments(loans: list[Loan]) -> list[Repayment]:
        """
        From provided loans, generate corresponding repayments with slices and dates
        :param loans:
        :return:
        """

        repayments_repository: dict[str, Repayment] = {}

        for loan in loans:
            repayment = repayments_repository.get(loan.loanee.ID)

            if not repayment:
                repayment = Repayment(loan.loanee)
                repayments_repository[loan.loanee.ID] = repayment

            repayment.add_slices(loan.repayment_logic, loan.date)

        return list(repayments_repository.values())
