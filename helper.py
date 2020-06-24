import pandas as pd
import difflib
from typing import List
from numpy import nan, array

from tornado.web import RequestHandler, Application
from tornado.ioloop import IOLoop
from tornado.httpserver import HTTPServer


class ExcelHelper:
    """A simple excel helper that helps you modifies excel files easily"""

    def __init__(self, excel_file: str) -> None:
        self.filename = excel_file
        self.frame_data = pd.read_excel(excel_file, None)

    def add_column(self, name: str, values: List[str] = [], sheet: \
     str = 'Sheet1') -> "self":
        """
        This function add a column to the current excel sheet.
        Name: The new column name.
        Values: The new values of the new column.
        Sheet: The excel sheet to operate.
        """
        self.frame_data[sheet][name] = values if values else [
            nan]*len(self.frame_data[sheet])
        return self

    def insert_row(self, row_data: List[str] = None, sheet='Sheet1'):
        row_data = row_data if row_data else nan*len(self.frame_data[sheet])
        self.frame_data[sheet] = self.frame_data[sheet].append([row_data], ignore_index=True)
        return self

    def save_as(self, excel_file: str = None) -> None:
        excel_file = excel_file if excel_file is not None else self.filename
        writer = pd.ExcelWriter(self.filename)
        for k, v in self.frame_data.items():
            v.to_excel(writer, sheet_name=k, index=False)


class IndexHandler(RequestHandler):
    def get(self):
        self.write('hi')