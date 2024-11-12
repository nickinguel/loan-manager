import datetime
from models.loanee import Loanee
from utilities.misc_utility import MiscUtility


class Repayment:

    def __init__(self, loanee: Loanee):
        self.loanee = loanee
        self.slices: list[RepaymentSlice] = []

    def add_slices(self, logic: str, start_date: datetime.date):
        slice_amounts = [int(amount.strip()) for amount in logic.split("+")]

        next_date: datetime.date

        if len(self.slices) > 0 and (last_date := self.slices[-1].date) > start_date:
            next_date = last_date
        else:
            next_date = start_date

        for slice_amount in slice_amounts:
            next_date = MiscUtility.get_next_month(next_date)

            self.slices.append(RepaymentSlice(slice_amount, next_date))


class RepaymentSlice:
    def __init__(self, amount, date):
        self.date: datetime.date = date
        self.amount = amount

    # def __init__(self, mon¸th_number, year, amount):
    #     self.month_number = month_number
        # self.year = year
        # self.amount = amount
        # self.month = ConstData.months[month_number - 1]
