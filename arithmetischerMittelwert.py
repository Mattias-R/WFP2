# importing openpyxl module
import openpyxl

# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_Neuhaus.xlsx'

# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)
sheet_obj = newWorkbook.active
# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["Tabelle1"]
maxLines = sheet_obj.max_row
# Passing the column index to the worksheet and traversing through the each row of the column
i = 2

arr = []
anomalie = []
helper = []
test = 0
verteilung = []
tableName = firstWorksheet["C"]
warning = []
excelrange = ["C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","F","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AF","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BF","BW","BX","BY"]
for device in excelrange:
    tableName = firstWorksheet[device]
    for i in range(maxLines):
        if(i==0):
            i=1
            if (tableName[i].value == "NA"):
                i = i + 1
                continue
            arr.append(float(tableName[i].value))
            helper.append(float(tableName[i].value))
        if(tableName[i].value != "NA" and float(tableName[i].value) > 3000):
             tableName[i].value = "NA"
        if(tableName[i].value == "NA"):
            i = i + 1
            continue
        #add value in arr
        arr.append(float(tableName[i].value))
        #add value in helper
        helper.append(float(tableName[i].value))
        #subtract current value with last value to get the actual consumption
        anomalie.append(round((float(tableName[i].value) - helper[len(helper) -2]),2))
        #mittlere absolute Abweichung berechnen
        # https://welt-der-bwl.de/Mittlere-absolute-Abweichung#:~:text=Die%20mittlere%20absolute%20Abweichung%20vom%20Median%20ist%3A%20(%20%7C%201%2D,Standardabweichung.
        median = round(sum(anomalie) / len(anomalie), 2)
        #print(median)
        abweichung = 0
        abweich = []
        for k in anomalie:
            #print("k - median", k, median)
            abweichung = abs(k - median)
            abweich.append(abweichung)
    #    print("abweichung:", round(abweichung, 2))
    #    print("median: ", median)
        verteilung.append(round(abweichung, 2))
        abweichung = sum(abweich) / len(anomalie)
        if(anomalie[len(anomalie)-1] > median+anomalie[len(anomalie)-2]):
            if(tableName[i-1].value == "NA"):
                continue
            var = float(tableName[i].value) - float(tableName[i-1].value)
            if(var < 0.2):
                continue
            x = [round(abweichung,2), i, tableName[i-1].value, tableName[i].value, round(var,2)]
            warning.append(x)
            if(anomalie[len(anomalie)-1] == 0):
                warning.pop()

        i = i + 1

#am Ende wird folgendes Format ausgegeben: die akzeptable Abweichung. Die Zeile wo es passiert. der Wert am vortag und der Wert am Tag
    print("Spalte: "+ str(device) + " Errors: " + str(len(warning)) + " List: " + str(warning))
    warning.clear()
    anomalie.clear()
    arr.clear()
    helper.clear()
#Probleme: urlaub... wenn lange kein Wasser verbraucht wird ist eine Meldung vorprogrammiert