
import pathlib
from pyspreadsheet import ExcelReader

# Read a excel sheet
DIRECTORY_OF_THIS_FILE = pathlib.Path(__file__).parent
excel = ExcelReader(DIRECTORY_OF_THIS_FILE / 'pyspreadsheet_test.xlsx')

# Access
table = excel['Equipment']
for row in table.rows:
    equipment = row['Type']
    model = row['Model'] 
    print(f' {equipment}: {model}')
    print(f'   row.reference: {row.reference}')
    print(f'   equipment.reference: {equipment.reference}')

# Dump the whole file for source code revision control
FILENAME_DUMP = 'pyspreadsheet_dump.txt'
print(f'Write: {FILENAME_DUMP}')
with (DIRECTORY_OF_THIS_FILE / FILENAME_DUMP).open('w') as f:
    excel.dump(f)
