import math as _math
import os as _os

import lib_dzne_filedata as _fd
import openpyxl as _xl
import pandas as _pd


class WorkbookData(_fd.FileData):
    _ext = '.xlsx'
    def __init__(self, workbook):
        self.workbook = workbook
    @classmethod
    def _load(cls, /, file):
        return _xl.load_workbook(file)
    def _save(self, /, file):
        self._workbook.save(filename=file)
    @staticmethod
    def _default():
        return _xl.Workbook()
    @classmethod
    def clone_data(workbook):
        with _tmp.TemporaryDirectory() as directory:
            file = _os.path.join(directory, "a" + cls.ext())
            workbook.save(file)
            return _xl.load_workbook(file)
    @staticmethod
    def workbook_from_DataFrames(dataFrames):
        dataFrames = dict(dataFrames)
        if len(dataFrames) == 0:
            return None
        workbook = _xl.Workbook()
        default_sheet = workbook.active
        for table, df in dataFrames.items():
            if default_sheet is None:
                workbook.create_sheet(table)
            else:
                default_sheet.title = table
                default_sheet = None
        for table, df in dataFrames.items():
            columns = list(df.columns)
            for x, column in enumerate(columns):
                workbook[table].cell(row=1, column=x+1).value = column
                for y, v in enumerate(df[column].tolist()):
                    if _pd.isna(v):
                        continue
                    elif (type(v) is float) and (_math.isinf(v)):# is this really needed?
                        value = str(v)
                    else:
                        value = v
                    workbook[table].cell(row=y+2, column=x+1).value = value
        return workbook
    @staticmethod
    def set_cell(*, cell, value):
        """Setting value of cell. """
        if _pd.isna(value):
            value = 'N/A'
        else:
            if type(value) is float:
                if _math.isinf(value):
                    if value < 0:
                        value = '-inf'
                    else:
                        value = '+inf'
            if type(value) not in {str, int, float, bool}:
                raise TypeError(f"The value {value} is of the invalid type {type(value)}! ")
        cell.value = value
        cell.alignment = _xl.styles.Alignment()#horizontal='general')



