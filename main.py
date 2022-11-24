# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\data.xlsx'

# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)

# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["data"]

# Passing the column index to the worksheet and traversing through the each row of the column
i = 0
datei = open(r'C:\Users\matti\PycharmProjects\csvIntoJson\sendData.txt', 'a')
datei.write("\r\n")
for column_data in firstWorksheet['D']:
   # Printing the column values of every row
   test = firstWorksheet['B']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":",column_data.value,",", end='')
   print("\"values\":{\"value\":",test, "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write("}}\" localhost:8080/api/v1/wuwitVXr9BqqDLxjHmlc/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
datei.close()
#