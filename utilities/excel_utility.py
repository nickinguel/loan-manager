from configs.constant_data import ConstData


class ExcelUtility:

    @staticmethod
    def get_repayment_cell_from_column_name(column_name: str, data_row_number) -> str | None:
        """
        Given a column name and corresponding data row (ignoring headers row), retrieve the corresponding cell index suc as A1, C4

        :param column_name:
        :param data_row_number:
        :return:
        """

        if column_name not in ConstData.excel_cols_repayments:
            return

        index = "{0}{1}".format(
            ConstData.alphabet[ConstData.excel_cols_repayments.index(column_name)],
            data_row_number + 1
        )

        return index
