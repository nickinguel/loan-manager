import sys

from configs.configs import Configs
from configs.constant_data import ConstData
from models.loan import Loan
from utilities.loan_utility import LoanUtility
from utilities.log_utility import LogUtility
from utilities.misc_utility import MiscUtility
from utilities.repayment_utility import RepaymentUtility


def main():

    file_path: str | None = None
    # file_path = '/Volumes/MyData/Temp/Harelle/Loans.xlsx'
    # file_path = 'D:\Temp\Harelle/Loans.xlsx'

    Configs.app_is_launched_on_console_mode = True

    if len(sys.argv) >= 2:
        file_path = sys.argv[1]

    if file_path is None or len(file_path.strip()) == 0:
        file_path = MiscUtility.browse_file_path("Choisissez le fichier Excel à manipuler")
        Configs.app_is_launched_on_console_mode = False

    loans: list[Loan]
    no_error = True

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

        # Write data to workbook
        RepaymentUtility.write_repayments_to_excel(file_path, sheets_data)

        # Find paid slices
        paid_slices = RepaymentUtility.find_paid_slices(file_path)

        # Compute stats data
        stats_data = RepaymentUtility.compute_stats(loans, paid_slices)

        # Write stats
        RepaymentUtility.write_stats_to_excel(stats_data, file_path)


    except Exception as e:
        LogUtility.log_error(e)
        no_error = False
        # raise e

    if no_error:
        LogUtility.log_success(ConstData.message_all_ok.format(file_path))

    if Configs.app_is_launched_on_console_mode:
        print()
        LogUtility.log_info(LogUtility.format_brand_message())


if __name__ == '__main__':
    main()

