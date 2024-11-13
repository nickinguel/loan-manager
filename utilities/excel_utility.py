from configs.constant_data import ConstData


class ExcelUtility:

    @staticmethod
    def get_cell_from_column_name(column_name: str, data_row_number, columns: tuple[str, ...] = ConstData.excel_cols_repayments) -> str | None:
        """
        Given a column name and corresponding data row (ignoring headers row), retrieve the corresponding cell index suc as A1, C4

        :param column_name:
        :param data_row_number:
        :param columns
        :return:
        """

        if column_name not in columns:
            return

        index = "{0}{1}".format(
            ConstData.alphabet[columns.index(column_name)],
            data_row_number + 1
        )

        return index
