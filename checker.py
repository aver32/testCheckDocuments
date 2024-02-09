import re
import pandas as pd

from constants import COLUMNS_NAME_FOR_READER_DF, INPUT_DATA_FILENAME, TABLES_DATA_FOLDER_NAME, OK, NOT_FOUND
from reader import DataReader


class DataChecker:
    def __init__(self):
        self.__reader = DataReader(INPUT_DATA_FILENAME, TABLES_DATA_FOLDER_NAME)
        self.__reader.read_documents()
        self.__result_df: pd.DataFrame = pd.DataFrame(columns=COLUMNS_NAME_FOR_READER_DF)

    @property
    def result_df(self):
        return self.__result_df

    def check_documents(self):
        self.__fill_input_mark_and_input_value()
        marks = self.__result_df[COLUMNS_NAME_FOR_READER_DF[0]].values
        for index_row, mark in enumerate(marks):
            value_in_register = self.__get_value_in_register_by_mark(mark)
            total_value_in_journal = self.__get_total_value_in_journal_by_mark(mark)
            status_res_inf = self.__get_status_res_inf_by_mark(mark)
            status_incheck = self.__get_status_incheck_by_mark(mark)
            self.__result_df.at[index_row, COLUMNS_NAME_FOR_READER_DF[2]] = value_in_register
            self.__result_df.at[index_row, COLUMNS_NAME_FOR_READER_DF[3]] = total_value_in_journal
            self.__result_df.at[index_row, COLUMNS_NAME_FOR_READER_DF[4]] = status_res_inf
            self.__result_df.at[index_row, COLUMNS_NAME_FOR_READER_DF[5]] = status_incheck

    def __get_status_res_inf_by_mark(self, mark):
        status_res_inf = (self.__reader.res_inf_df.iloc[:, 1] == mark).any()
        if status_res_inf:
            return OK
        else:
            return NOT_FOUND

    def __get_status_incheck_by_mark(self, mark):
        status_incheck = self.__reader.incheck_df[(self.__reader.incheck_df.iloc[:, 1].str.contains(mark)) |
                                                  (self.__reader.incheck_df.iloc[:, 9].str.contains(mark)) |
                                                  (self.__reader.incheck_df.iloc[:, 10].str.contains(mark))].iloc[:, 17]
        if not status_incheck.empty:
            return f"{OK}, {status_incheck.values[0].strip()}"
        else:
            return NOT_FOUND

    def __get_total_value_in_journal_by_mark(self, mark: str):
        total_value_in_journal = self.__reader.journal_df.loc[
                                     self.__reader.journal_df.iloc[:, 2].str.contains(mark)].iloc[:, 3].sum()
        return total_value_in_journal

    def __get_value_in_register_by_mark(self, mark: str) -> int:
        value_in_register = self.__reader.register_df.loc[self.__reader.register_df.iloc[:, 2] == mark].iloc[:, 4]
        if not value_in_register.empty:
            meters = re.search(r"\d+", value_in_register.values[0])[0]
            return int(meters)
        else:
            return 0

    def __fill_input_mark_and_input_value(self):
        for index_column, column_name in enumerate(COLUMNS_NAME_FOR_READER_DF[:2]):
            self.__result_df[column_name] = self.__reader.input_data_df.iloc[:, index_column].values
