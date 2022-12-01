# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\test.xlsx'

# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)
sheet_obj = newWorkbook.active
print(sheet_obj.max_row)

# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["test"]

# Passing the column index to the worksheet and traversing through the each row of the column
i = 0
datei = open(r'C:\Users\matti\PycharmProjects\csvIntoJson\sendData.txt', 'a')
datei.write("\r\n")

#-----------------------------------------------------------------------
# Damit kann der token für jedes erstellte Gerät eingefügt und durchgearbetiet werden. Die meiste arbeit ist nur noch das Gerät zu erstellen
token = ["test1", "test2", "test3", "test4","test5","test6","test7"]
# Ascii -> A bis Z ist 65 bis 90. anfangen tun wir mit B, also 66
table = 66
#coutner mitzählen lassen, wie viel NA zwischen dem letzten NA und dem aktuellen wert waren.
#danach eventuell soviele zurückspringen, die differenz durch die anzahl an NAs und das jeden tag dazu addieren und die daten damit auffüllen.
#Mit backtrack alles durchgehen und zurückspringen?
#Ist wasserbilanz der positive wert und die netnahme dann die dinge die subtrahiert werden?

for device in token:
   for column_data in firstWorksheet['J']:
      # Printing the column values of every row
      print(""+chr(table))
      test = firstWorksheet[""+chr(table)]
      test = test[i].value
      print("sudo curl -v -X POST --data \"{\"ts\":",column_data.value,",", end='')
      print("\"values\":{\"value\":",test, "}}\" localhost:8080/api/v1/",device, "/telemetry --header \"Content-Type:application/json\"")
      datei.write("sudo curl -v -X POST --data \"{\"ts\":")
      datei.write(str(column_data.value))
      datei.write(",\"values\":{\"value\":")
      datei.write(str(test))
      datei.write("}}\" localhost:8080/api/v1/")
      datei.write(device)
      datei.write("/telemetry --header \"Content-Type:application/json\"\r\n")
      i = i + 1
   table = table +1
   i = 0
datei.close()
#