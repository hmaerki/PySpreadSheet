'''
This is an example on how to use 'pyspreadsheet'.

Below are doctests. It is worth to study them as they show benefits of this library.
'''

from pathlib import Path
from pyspreadsheet import ExcelReader

# Read a excel sheet
DIRECTORY_OF_THIS_FILE = Path(__file__).parent
excel = ExcelReader(DIRECTORY_OF_THIS_FILE / 'pyspreadsheet_test.xlsx')

from enum import Enum
class EquipmentType(Enum):
    Voltmeter = 1
    Multimeter = 2

print('\nAccess using indices:')
equipment = excel['Equipment']
for row in equipment.rows:
    instrument = row['Instrument']
    model = row['Model'] 
    print(f'  {instrument}: {model}')


print('\nAccess using properties:')
for row in excel.tables.Equipment.rows:
    print(f'  Instrument: "{row.cols.Instrument}" Model: "{row.cols.Model}"')
    print(f'    reference: {row.cols.Model.reference}')

# Dump the whole file for source code revision control
FILENAME_DUMP = 'pyspreadsheet_dump.txt'
print(f'\nWrite: {FILENAME_DUMP}')
with (DIRECTORY_OF_THIS_FILE / FILENAME_DUMP).open('w') as f:
    excel.dump(f)

def doctest_reference():
    '''
    >>> excel.reference
    'File "pyspreadsheet_test.xlsx"'

    >>> excel.tables.Equipment.reference
    'Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'

    >>> excel.tables.Equipment.rows[0].reference
    'Row 3 in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'

    >>> excel.tables.Equipment.rows[0].cols.ID.reference
    'Cell "C3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'
    '''

def doctest_excelreader():
    '''
    >>> ExcelReader('invalid_filename.xlsx')
    Traceback (most recent call last):
       ...
    FileNotFoundError: [Errno 2] No such file or directory: 'invalid_filename.xlsx'

    >>> excel.tables.Invalid
    Traceback (most recent call last):
       ...
    KeyError: 'No table "Invalid". Valid tables are "Equipment|Measurement|TableA|TableC". See: File "pyspreadsheet_test.xlsx"'

    >>> excel['Invalid']
    Traceback (most recent call last):
       ...
    KeyError: 'No table "Invalid". Valid tables are "Equipment|Measurement|TableA|TableC". See: File "pyspreadsheet_test.xlsx"'
    '''

def doctest_row():
    '''
    >>> row = excel.tables.Equipment.rows[0]

    >>> row['Invalid']
    Traceback (most recent call last):
       ...
    KeyError: 'No column "Invalid". Valid columns are ID|Instrument|Model|Serial. See: Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'

    >>> row.cols.Invalid
    Traceback (most recent call last):
       ...
    KeyError: 'No column "Invalid". Valid columns are ID|Instrument|Model|Serial. See: Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'
    '''

def doctest_cell_int():
    '''
    >>> cell_id = excel.tables.Equipment.rows[0].cols.ID

    >>> cell_id.reference
    'Cell "C3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"'

    >>> cell_id.text
    '1'

    >>> cell_id.int
    1

    >>> cell_id.astype(float)
    1.0

    >>> cell_voltmeter = excel.tables.Equipment.rows[0].cols.Instrument
    >>> cell_voltmeter.int
    Traceback (most recent call last):
       ...
    ValueError: "Voltmeter" is not a valid int! See: Cell "D3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"
    '''

def doctest_cell_enumaration():
    '''
    >>> cell_voltmeter = excel.tables.Equipment.rows[0].cols.Instrument

    >>> cell_voltmeter.text
    'Voltmeter'

    >>> cell_voltmeter.asenum(EquipmentType)
    <EquipmentType.Voltmeter: 1>

    >>> cell_id = excel.tables.Equipment.rows[0].cols.ID
    >>> cell_id.asenum(EquipmentType)
    Traceback (most recent call last):
       ...
    ValueError: "1" is not a valid EquipmentType! Valid values are Voltmeter|Multimeter. See: Cell "C3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"
    '''

def doctest_assert_not_empty():
    '''
    >>> cell_empty = excel.tables.Equipment.rows[0].cols.Serial

    >>> cell_empty.text_not_empty
    Traceback (most recent call last):
       ...
    Exception: Cell must not be empty! Cell "F3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"
    
    >>> cell_voltmeter = excel.tables.Equipment.rows[0].cols.Instrument
    >>> cell_voltmeter.text_not_empty
    'Voltmeter'
    '''

def doctest_date():
    '''
    >>> cell_date = excel.tables.Measurement.rows[0].cols.Date
    >>> cell_date.text
    '2017-06-20 00:00:00'
    >>> cell_date.asdate()
    '2017-06-20'
    >>> cell_date.asdate(format='%A')
    'Tuesday'
    >>> cell_voltmeter = excel.tables.Equipment.rows[0].cols.Instrument
    >>> cell_voltmeter.asdate()
    Traceback (most recent call last):
       ...
    ValueError: "Voltmeter" is not a datetime! See: Cell "D3" in Table "Equipment" in Worksheet "Inventory" in File "pyspreadsheet_test.xlsx"
    '''

if __name__ == '__main__':
    import doctest
    doctest.testmod()
