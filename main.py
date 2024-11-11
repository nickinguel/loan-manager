from models.loan import Loan
from utilities.loan_utility import LoanUtility
from utilities.log_utility import LogUtility


def main():
    file_path = '/Volumes/MyData/Temp/Harelle/Loans.xlsx'

    loans: list[Loan]

    try:
        LoanUtility.read_loans(file_path)
    except Exception as e:
        LogUtility.log_error(e)
        raise e


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
