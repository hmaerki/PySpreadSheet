# PySpreadSheet
Read data from Excel SpreadSheets. Based on openpyxl.

This is a python implementation of a subset of https://github.com/hmaerki/ZuluSpreadSheet.

## Example Excel Sheet

`pyspreadsheet_test.xlsx`
![Kiku](images/tables.png)

## Example Code

```python
excel = ExcelReader('zulu_excel_reader_test.xlsx')
table = excel['Equipment']
for row in table.rows:
    equipment = row['Type']
    model = row['Model'] 
    print(f' {equipment}: {model}')
    print(f'   row.reference: {row.reference}')
    print(f'   equipment.reference: {equipment.reference}')

excel.dump('zulu_excel_reader_test_dump.txt')
```

### Output

```text
Voltmeter: Keysight 34460A
  row.reference: Row 3 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
  equipment.reference: Cell D3 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
Multimeter: Fluke 787
  row.reference: Row 4 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
  equipment.reference: Cell D4 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
Voltmeter: KEYSIGHT U1231A
  row.reference: Row 5 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
  equipment.reference: Cell D5 in Table "SheetQuery" in Worksheet "Equipment" in File "zulu_excel_reader_test.xlsx""
```

### How to follow changes in a binary excel sheet?

The dump-file may be used to track relevant changes in the excel file. Specially in git it is important to follow the changes in the content of the excel sheet.

`pyspreadsheet_dump.txt`
```text
Table: Equipment
  ID|Model|Serial|Type

  1|Keysight 34460A|1245678|Voltmeter
  2|Fluke 787|2234|Multimeter
  3|KEYSIGHT U1231A|134555|Voltmeter

Table: Measurement
  Date|Equipment|Operator

  42906|1|Karl
  42906|3|Rosa
  42898|1|Otto
  42891|2|Karl
```

