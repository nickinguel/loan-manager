from models.loan import Loan


class RepaymentUtility:

    @staticmethod
    def convert_loan_to_sheet_data(loans: list[Loan]) -> dict[str, list[list]]:
        sheets_data = {}

        for loan in loans:
            pass

        return sheets_data