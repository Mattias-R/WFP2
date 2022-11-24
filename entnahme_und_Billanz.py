# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\test.xlsx'

# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)

# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["test"]

# Passing the column index to the worksheet and traversing through the each row of the column
i = 0
datei = open(r'C:\Users\matti\PycharmProjects\csvIntoJson\sendData.txt', 'a')
datei.write("\r\n")
#QS Neuhaus
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['B']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":",column_data.value,",", end='')
   print("\"values\":{\"value\":",test, "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write("}}\" localhost:8080/api/v1/vVWuOI2UlWEcp4YovRxK/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
#HB Neuhaus Entnahme
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['C']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":",column_data.value,",", end='')
   print("\"values\":{\"value\":",test, "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write("}}\" localhost:8080/api/v1/D99mQtr7AqJn0XJeawjA/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
# HB Neuhaus Zulauf
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['D']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":", column_data.value, ",", end='')
   print("\"values\":{\"value\":", test,
         "}}\" localhost:8080/api/v1/bKM9fpZcfJmQ4fDZ8Lwe/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write(
      "}}\" localhost:8080/api/v1/bKM9fpZcfJmQ4fDZ8Lwe/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
# HB Neuhaus Verbrauch
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['E']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":", column_data.value, ",", end='')
   print("\"values\":{\"value\":", test,
         "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write(
      "}}\" localhost:8080/api/v1/jovF1SXv0xpi1V5VP7Kn/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
# HB Pudlach Entnahme
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['F']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":", column_data.value, ",", end='')
   print("\"values\":{\"value\":", test,
         "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write(
      "}}\" localhost:8080/api/v1/MkOFhnKPMkORI22OBbgx/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
# HB Schwabegg Entnahme
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['G']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":", column_data.value, ",", end='')
   print("\"values\":{\"value\":", test,
         "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write(
      "}}\" localhost:8080/api/v1/YjBxQeYoObAklNc3IE8O/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
i = 0
# HB HB Neuhaus Wasserbilanz
for column_data in firstWorksheet['J']:
   # Printing the column values of every row
   test = firstWorksheet['G']
   test = test[i].value
   print("sudo curl -v -X POST --data \"{\"ts\":", column_data.value, ",", end='')
   print("\"values\":{\"value\":", test,
         "}}\" localhost:8080/api/v1/fxsKjElrtIRpx7r2qk35/telemetry --header \"Content-Type:application/json\"")
   datei.write("sudo curl -v -X POST --data \"{\"ts\":")
   datei.write(str(column_data.value))
   datei.write(",\"values\":{\"value\":")
   datei.write(str(test))
   datei.write(
      "}}\" localhost:8080/api/v1/MrxGd20EuinPBoJdlbwP/telemetry --header \"Content-Type:application/json\"\r\n")
   i = i + 1
datei.close()