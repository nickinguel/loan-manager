import datetime
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet._reader import Cell
from openpyxl.worksheet.worksheet import Worksheet
from configs.constant_data import ConstData
from models.loan import Loan
from models.loanee import Loanee
from utilities.misc_utility import MiscUtility
import os
import re


class LoanUtility:

    @staticmethod
    def read_loans(file_path: str) -> list[Loan]:
        workbook: Workbook = None
        loans: list[Loan] = []

        try:
            if not os.path.exists(file_path):
                raise Exception(f"Aucun fichier trouvé à l'emplacement spécifié '{file_path}'")

            workbook = MiscUtility.open_workbook(file_path)
            loan_sheet = workbook.get_sheet_by_name(ConstData.excel_sheet_loan)

            # Check missing headers
            header_tuple = LoanUtility.check_required_columns_headers(loan_sheet)
            headers = header_tuple[0]
            missing_headers = header_tuple[1]

            if len(missing_headers) > 0:
                raise Exception("Dans la feuille '{0}', les colonnes suivantes sont manquantes : {1}".format(
                    ConstData.excel_sheet_loan,
                    MiscUtility.format_array_as_bullets(missing_headers)
                ))

            # Parse loans
            loans = LoanUtility.parse_loans(loan_sheet, headers)

        except Exception as e:
            raise e
        finally:
            if workbook is not None:
                MiscUtility.close_workbook(workbook)

        return loans

    @staticmethod
    def parse_loans(sheet: Worksheet, headers: dict[str, int]) -> list[Loan]:
        """
        Parse the Loan sheet and convert rows to corresponding :Loan object
        :param sheet:
        :param headers:
        :return:
        """

        loans: list[Loan] = None

        for cells_tpl in sheet.iter_rows(min_row=2, values_only=False):
            cells_values = tuple([cell.value for cell in cells_tpl])
            missing_values = LoanUtility.check_loan_row(headers, cells_tpl)

            if len(missing_values) > 0:
                raise Exception("Dans la feuille '{0}', la ligne '{1}' a des valeurs manquantes ou incorrectes pour les colonnes suivantes : {2}"
                                .format(ConstData.excel_sheet_loan, cells_tpl[0].row.real, MiscUtility.format_array_as_bullets(missing_values))
                            )
            loanee = Loanee(
                cells_values[headers[ConstData.excel_col_loan_debtor_ID]],
                cells_values[headers[ConstData.excel_col_loan_debtor_first_name]],
                cells_values[headers[ConstData.excel_col_loan_debtor_last_name]]
            )

            # loans.append(Loan())

        return loans

    @staticmethod
    def check_loan_row(headers: dict[str, int], row_values: tuple[Cell, ...]) -> list[str]:
        """
        Check if some data rows have wrong values (empty or incorrect type)
        :param headers: Headers column indexes
        :param row_values: The values in the row for each column
        :return:
        """

        missing_cells = []
        headers_copy = dict((key, val) for key, val in headers.items() if key not in [
            ConstData.excel_col_loan_amount,
            ConstData.excel_col_loan_date,
            ConstData.excel_col_loan_repayment_logic
        ])

        if re.search("^\\d+$", str(row_values[headers[ConstData.excel_col_loan_amount]].value)) is None:
            missing_cells.append(ConstData.excel_col_loan_amount)

        if re.search(r"^\d+(\s*[+*]\s*\d+)*$", str(row_values[headers[ConstData.excel_col_loan_repayment_logic]].value)) is None:
            missing_cells.append(ConstData.excel_col_loan_repayment_logic)
        else:
            converted_logic = LoanUtility.convert_repayment_logic(
                row_values[headers[ConstData.excel_col_loan_repayment_logic]].value,
                int(row_values[headers[ConstData.excel_col_loan_amount]].value)
            )
            repayment_amount = eval(converted_logic)

            print(str(row_values[headers[ConstData.excel_col_loan_repayment_logic]].value), " --> ", converted_logic)

            if int(row_values[headers[ConstData.excel_col_loan_amount]].value) != repayment_amount:
                missing_cells.append(ConstData.excel_col_loan_repayment_logic)

        try:
            _ = from_excel(row_values[headers[ConstData.excel_col_loan_date]].value)
        except (Exception,):
            missing_cells.append(ConstData.excel_col_loan_date)

        for col_name, col_index in headers_copy.items():
            if (cell_val := row_values[col_index].value) is None or len(str(cell_val)) == 0:
                missing_cells.append(col_name)

        return missing_cells

    @staticmethod
    def convert_repayment_logic(logic: str, loan_amount: int):
        """
        Takes the original repayment logic from the excel file and convert it to basic + operation
        Exemple : 1000 * 2 ==> 1000 + 1000
        :param logic:
        :param loan_amount:
        :return:
        """

        logic_str = str(logic)

        if re.search(r"^\d+$", logic_str) is not None:
            slice_amount = int(logic_str)
            return LoanUtility.write_slice_n_times(loan_amount // slice_amount, slice_amount)

        while (re_match := re.search(r"(\d+)\s*\*\s*(\d+)", logic_str)) is not None:
            times: int
            amount: int
            multiplication = re_match.group()
            groups = re_match.groups()

            if (first := int(groups[0])) < (second := int(groups[1])):
                times = first
                amount = second
            else:
                times = second
                amount = first

            logic_str = logic_str.replace(multiplication, LoanUtility.write_slice_n_times(times, amount), 1)

        return logic_str

    @staticmethod
    def write_slice_n_times(times: int, amount: any):
        return " + ".join([str(amount) for i in range(times)])

    @staticmethod
    def check_required_columns_headers(sheet: Worksheet):
        """
        Check if some required column headers are missing in the Excel file
        :param sheet:
        :return:
        """
        loan_col_headers = [cell.value for cell in sheet[1]]

        missing_headers = [h for h in ConstData.excel_cols_loan if h not in loan_col_headers]
        headers = dict([(cell.value, index) for index, cell in enumerate(sheet[1]) if cell.value in ConstData.excel_cols_loan])

        return headers, missing_headers

