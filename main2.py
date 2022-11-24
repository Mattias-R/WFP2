# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\daten.xlsx'

# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)

# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["daten"]

# Passing the column index to the worksheet and traversing through the each row of the column
for column_data in firstWorksheet['B']:
   # Printing the column values of every row
   print("\"values\":{\"value\":",column_data.value, "}}\" localhost:8080/api/v1/PWWRbiayfsNsY9NrqjQx/telemetry --header \"Content-Type:application/json\""
)
