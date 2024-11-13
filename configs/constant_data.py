import datetime


class ConstData:

    # Misc
    months = [datetime.date(1990, month, 1).strftime("%B") for month in range(1, 13)]
    alphabet = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')

    # Loans
    excel_sheet_loan = "Emprunts"

    excel_col_loan_debtor_ID = "ID"
    excel_col_loan_debtor_first_name = "Prenoms"
    excel_col_loan_debtor_last_name = "Nom"
    excel_col_loan_amount = "Montant"
    excel_col_loan_repayment_logic = "Remboursement"
    excel_col_loan_date = "Date"

    excel_cols_loan = (excel_col_loan_debtor_ID, excel_col_loan_debtor_first_name, excel_col_loan_debtor_last_name, excel_col_loan_amount,
                       excel_col_loan_date, excel_col_loan_repayment_logic)

    # Repayments
    excel_sheet_repayment = "Remboursements"
    excel_col_repayment_total_amount_loaned = "Total annuel emprunté"
    excel_cols_repayments = (excel_col_loan_debtor_ID, excel_col_loan_debtor_first_name, excel_col_loan_debtor_last_name,
                             excel_col_repayment_total_amount_loaned, *months)

    # Messages
    message_all_ok = "Les feuilles de remboursements ont été générées avec succès dans le fichier '{0}'"
    message_loan_sheet_missing = "La feuille '{0}'".format(excel_sheet_loan) + " n'existe pas dns le fichier spécifié '{0}'"

    # Stats
    excel_col_stats_loan_total = "Total emprunté"
    excel_col_stats_loan_total_refunded = "Total remboursé"
    excel_col_stats_loan_total_remaining = "Total prêt restant"
    excel_cols_stats = (excel_col_loan_debtor_ID, excel_col_loan_debtor_first_name, excel_col_loan_debtor_last_name, excel_col_stats_loan_total,
                        excel_col_stats_loan_total_refunded, excel_col_stats_loan_total_remaining)
