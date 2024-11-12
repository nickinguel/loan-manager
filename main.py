from models.loan import Loan
from utilities.loan_utility import LoanUtility
from utilities.log_utility import LogUtility
from datetime import datetime


def main():
    file_path = '/Volumes/MyData/Temp/Harelle/Loans.xlsx'

    loans: list[Loan]

    try:
        loans = LoanUtility.read_loans(file_path)
        loans.sort(key=lambda p: f"{p.loanee.ID}-{p.date.strftime("%m/%d/%Y")}", reverse=True)

        print([(loan.amount, loan.loanee.first_name, loan.date) for loan in loans])
    except Exception as e:
        LogUtility.log_error(e)
        raise e


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
