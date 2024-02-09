import sys
import pandas as pd

from constants import TABLE_REGISTER_FILENAME, TABLE_INCHECK_FILENAME, TABLE_JOURNAL_FILENAME, TABLE_RES_INF_FILENAME


class DataReader:
    def __init__(self, input_data_filename: str, table_data_folder_name: str):
        self.__filename = input_data_filename
        self.__table_data_folder_name = table_data_folder_name
        self.__input_data_df: pd.DataFrame or None = None
        self.__result_df: pd.DataFrame or None = None
        self.__register_df: pd.DataFrame or None = None
        self.__incheck_df: pd.DataFrame or None = None
        self.__journal_df: pd.DataFrame or None = None
        self.__res_inf_df: pd.DataFrame or None = None

    @property
    def input_data_df(self):
        return self.__input_data_df

    @property
    def register_df(self):
        return self.__register_df

    @property
    def incheck_df(self):
        return self.__incheck_df

    @property
    def journal_df(self):
        return self.__journal_df

    @property
    def res_inf_df(self):
        return self.__res_inf_df

    def read_documents(self):
        self.create_dataframes_from_files()

    def create_dataframes_from_files(self):
        try:
            self.__input_data_df = pd.read_excel(io=self.__filename, index_col=0)
            self.__register_df = pd.read_csv(
                filepath_or_buffer=f"{self.__table_data_folder_name}/{TABLE_REGISTER_FILENAME}",
                encoding='windows-1251',
                sep=';')
            self.__incheck_df = pd.read_csv(
                filepath_or_buffer=f"{self.__table_data_folder_name}/{TABLE_INCHECK_FILENAME}",
                encoding='windows-1251',
                sep=';')
            self.__journal_df = pd.read_csv(
                filepath_or_buffer=f"{self.__table_data_folder_name}/{TABLE_JOURNAL_FILENAME}",
                encoding='windows-1251',
                sep=';')
            self.__res_inf_df = pd.read_csv(
                filepath_or_buffer=f"{self.__table_data_folder_name}/{TABLE_RES_INF_FILENAME}",
                encoding='windows-1251',
                sep=';')
        except FileNotFoundError:
            print(f"Произошла ошибка при чтении файлов, скрипт закрывается")
            sys.exit(1)
