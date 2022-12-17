# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_Neuhaus.xlsx'
#inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_schwabbeg.xlsx'
# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)
sheet_obj = newWorkbook.active
# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["Tabelle1"]
maxLines = sheet_obj.max_row
# Passing the column index to the worksheet and traversing through the each row of the column
excelrange = ["C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","F","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AF","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BF","BW","BX","BY"]
test = ["C"]
count = 0

for device in excelrange:
    tableName = firstWorksheet[device]
    printErrors = []
    median = []
    value = []
    helper = []
    abweichung = 0
    for i in range(maxLines):
        #is needed to start
        if (i == 0):
            continue
        #if value is unrealistic, make it NA
        if(tableName[i].value != "NA" and float(tableName[i].value) > 3000):
            tableName[i].value = "NA"
        #skip loop if value is NA
        if(tableName[i].value == "NA"):
            continue
        if(str(tableName[i-1].value).startswith("WM")):
            continue
        helper.append(float(tableName[i].value))
        if(tableName[i-1].value == "NA"):
            difference = round(float(tableName[i].value) - float(helper[len(helper)-2]) ,2)
            value.append(difference)
            continue
        else:
            difference = round(float(tableName[i].value) - float(tableName[i-1].value) ,2)
            value.append(difference)
        if(value[len(value)-1] == 0):
            value.pop()
        median = round(float(sum(value) / float(len(value))),2)
        for k in value:
            abweichung = abweichung + abs(k - median)
        abweichung = round(abweichung / float(len(value)),2)

        if(difference > (median + abweichung)*2):
            x = [i , round(median+abweichung,2), difference]
            printErrors.append(x)
            count = count + 1

    print("Spalte: " + str(device) + " Errors: " + str(len(printErrors)) + " List: " + str(printErrors))
print("All errors: ",count)
