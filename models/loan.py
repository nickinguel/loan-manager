import datetime

from models.loanee import Loanee


class Loan:

    def __init__(self, loanee: Loanee, amount: int, repayment_logic: str, date: datetime.date):
        self.loanee = loanee
        self.amount = amount
        self.repayment_logic = repayment_logic
        self.date: datetime = date
