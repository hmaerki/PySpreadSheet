# PySpreadSheet
# https://github.com/hmaerki/PySpreadSheet
# (c) Copyright 2002-2020, Hans Maerki
# Distributed under GNU LESSER GENERAL PUBLIC LICENSE Version 3

# Documentation of openpyxl:
# https://bitbucket.org/openpyxl/openpyxl
# https://openpyxl.readthedocs.io/en/stable/
import pathlib
import openpyxl

TAG_TABLE = 'TABLE'
TAG_HYPHEN = '-'
TAG_EMPTY = ''
COLUMN_TABLE = 0 # TABLE
COLUMN_NAME = 1 # Table name
COLUMN_FIRST_CULUMN = 2

class ExcelReaderExecption(Exception):
    pass

class Cell:
    def __init__(self, row, table, idx):
        self.__row = row
        self.__table = table
        self.__idx = idx
        self.__cell = Cell.get_cell(row, idx)

    @property
    def text(self):
        return str(self.__cell.value)

    def __str__(self):
        return self.text

    @property
    def reference(self):
        return f'Cell "{self.__cell.coordinate}" in {self.__table.reference}'

    @classmethod
    def get_cell(cls, row, idx):
        if idx >= len(row):
            return None
        return row[idx]

    @classmethod
    def get_cell_value(cls, row, idx):
        cell = Cell.get_cell(row, idx)
        if cell is None:
            return ''
        if cell.value is None:
            return ''
        return str(cell.value)

class Row:
    def __init__(self, table, row, rowidx):
        self.__table = table
        self.__row = row
        self.__rowidx = rowidx

    @property
    def reference(self):
        return f'Row {self.__rowidx+1} in {self.__table.reference}'

    def __getitem__(self, column_name):
        if isinstance(column_name, str):
            idx = self.__table.get_columnidx_by_name(column_name)
        else:
            assert isinstance(column_name, int)
            idx = column_name
        return Cell(self.__row, self.__table, idx)

    def dump(self, file):
        columns = [str(self[c]) for c in self.__table.column_names]
        print(f'  {"|".join(columns)}', file=file)

class Table:
    def __init__(self, excel, table_name, worksheet_name):
        self.__excel = excel
        self.table_name = table_name
        self.worksheet_name = worksheet_name
        self.rows = []
        self.__columnname2idx = {}

    @property
    def reference(self):
        return f'Table "{self.table_name}" in Worksheet "{self.worksheet_name}" in {self.__excel.reference}'

    @property
    def column_names(self):
        return sorted(self.__columnname2idx.keys())

    @property
    def column_names_text(self):
        return '|'.join(self.column_names)

    def parse_columns(self, obj_row):
        for i in range(COLUMN_FIRST_CULUMN, len(obj_row)):
            column_name = Cell.get_cell_value(obj_row, i).strip()
            if column_name == TAG_HYPHEN:
                continue
            if column_name == TAG_EMPTY:
                break
            self.__columnname2idx[column_name] = i

    def get_columnidx_by_name(self, column_name):
        try:
            return self.__columnname2idx[column_name]
        except KeyError as e:
            raise KeyError(f'No column "{column_name}". Valid columns are "{self.column_names_text}".') from e

    def add_row(self, obj_row, rowidx):
        self.rows.append(Row(self, obj_row, rowidx))

    def dump(self, file):
        print(file=file)
        print(f'Table: {self.table_name}', file=file)
        print(f'  {self.column_names_text}', file=file)
        print(file=file)
        for obj_row in self.rows:
            obj_row.dump(file=file)

    def raise_exception(self, str_msg):
        raise ExcelReaderExecption(f'Table "{self.table_name}": {str_msg}')

class ExcelReader:
    def __init__(self, str_filename_xlsx):
        self.__filename = str_filename_xlsx
        self.__dict_tables = {}
        actual_table = None

        workbook = openpyxl.load_workbook(filename=str_filename_xlsx, read_only=True, data_only=True)
        for worksheet in workbook.worksheets:
            for rowidx, row in enumerate(worksheet.rows):
                if len(row) < COLUMN_FIRST_CULUMN:
                    actual_table = None
                    continue
                def get_value(row, idx):
                    val = row[idx].value
                    if val is None:
                        return ''
                    return str(val)
                cell_table = get_value(row, COLUMN_TABLE)
                cell_name = get_value(row, COLUMN_NAME)

                if actual_table:
                    if cell_table == TAG_HYPHEN:
                        continue
                    if cell_table == TAG_EMPTY:
                        actual_table = None
                        continue
                    actual_table.add_row(row, rowidx)
                    continue

                if cell_table == TAG_TABLE:
                    if len(row) <= COLUMN_FIRST_CULUMN:
                        actual_table.raise_exception(f'Need at least {COLUMN_FIRST_CULUMN+1} rows')
                    table_name = cell_name
                    assert table_name is not None
                    actual_table = Table(self, table_name, worksheet.title)
                    self.__dict_tables[table_name] = actual_table
                    actual_table.parse_columns(row)
                    continue

    @property
    def reference(self):
        return f'File "{self.__filename.name}"'

    @property
    def table_names(self):
        return sorted(self.__dict_tables.keys())

    @property
    def table_names_text(self):
        return '|'.join(self.table_names)

    def __getitem__(self, table_name):
        try:
            return self.__dict_tables[table_name]
        except KeyError as e:
            raise KeyError(f'No table "{table_name}". Valid tables are "{self.table_names_text}".') from e

    def dump(self, file):
        if isinstance(file, pathlib.Path):
            with file.open('w') as f:
                self.dump(f)
            return

        for table_name in self.table_names:
            self[table_name].dump(file)
