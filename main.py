from models.loan import Loan
from utilities.loan_utility import LoanUtility
from utilities.log_utility import LogUtility
from utilities.repayment_utility import RepaymentUtility


def main():
    file_path = '/Volumes/MyData/Temp/Harelle/Loans.xlsx'

    loans: list[Loan]

    try:
        # Retrieve and parse Loans from Excel
        loans = LoanUtility.read_loans(file_path)
        loans.sort(key=lambda loan: f"{loan.loanee.ID}-{loan.date.strftime("%m/%d/%Y")}", reverse=True)
        # print([(loan.amount, loan.loanee.first_name, loan.date) for loan in loans])
        # print()

        # Compute loan repayments
        repayments = RepaymentUtility.compute_repayments(loans)
        # print([(rep.loanee.ID, [(sl.amount, sl.date.strftime('%B %Y')) for sl in rep.slices]) for rep in repayments])
        # print()

        # Group repayments by year
        repayments_grouped = RepaymentUtility.group_repayments_by_year(repayments)

        # Compute sheets data
        sheets_data = RepaymentUtility.compute_sheets_data(repayments_grouped)
        print(sheets_data)

        # Write data to workbook
        RepaymentUtility.write_repayments_to_excel(file_path, sheets_data)

    except Exception as e:
        LogUtility.log_error(e)
        raise e


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
