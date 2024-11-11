from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from configs.constant_data import ConstData
from utilities.misc_utility import MiscUtility


class LoanUtility:

    @staticmethod
    def read_loans(file_path: str) -> list[list]:
        workbook = MiscUtility.open_workbook(file_path)
        loan_sheet = workbook.get_sheet_by_name(ConstData.excel_sheet_loan)

        missing_headers = LoanUtility.check_required_columns_headers(loan_sheet)

        if len(missing_headers) > 0:
            raise Exception("Dans la feuille '{0}', les colonnes suivantes sont manquantes : {1}".format(
                ConstData.excel_sheet_loan,
                MiscUtility.format_array_as_bullets(missing_headers)
            ))


        return None

    @staticmethod
    def check_required_columns_headers(sheet: Worksheet):
        loan_col_headers = [cell.value for cell in sheet[1]]
        missing_headers = [h for h in ConstData.excel_cols_loan if h not in loan_col_headers]

        return missing_headers

