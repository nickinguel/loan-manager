import datetime

from models.loanee import Loanee


class Loan:

    def __init__(self, loanee: Loanee, amount: int, repayment_logic: str, date: datetime.datetime):
        self.loanee = loanee
        self.amount = amount
        self.repayment_logic = repayment_logic
        self.date: datetime.date = date.date() if isinstance(date, datetime.datetime) else date
