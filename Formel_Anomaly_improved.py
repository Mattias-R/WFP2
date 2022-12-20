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
    durchschnitt = []
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
        #skip loop if value is WM
        if(str(tableName[i-1].value).startswith("WM")):
            continue
        #now only values should  be in tablename.value. therefore save it into the helper array
        helper.append(float(tableName[i].value))
        #if last value is NA than we have to use the last value from the helper array. therefore helper[len(helper)-2] because -1 would be the actual value which we added a second ago
        #if NA then continue, because the value would be too high. but we need the information for the caluclation.
        if(tableName[i-1].value == "NA"):
            difference = round(float(tableName[i].value) - float(helper[len(helper)-2]), 2)
            value.append(difference)
            continue
        else:
            difference = round(float(tableName[i].value) - float(tableName[i-1].value), 2)
            value.append(difference)
        # if value is 0, that means the customer is probably on holiday
        # if we do not pop it, the allowed difference would be too low and more false positive appear
        if(value[len(value)-1] == 0):
            value.pop()
        # calculates the durchschnitt
        durchschnitt = round(float(sum(value) / float(len(value))), 2)
        for k in value:
            # calculates the abweichung based on the formula of https://welt-der-bwl.de/Mittlere-absolute-Abweichung#:~:text=Die%20mittlere%20absolute%20Abweichung%20vom%20Median%20ist%3A%20(%20%7C%201%2D,Standardabweichung.
            abweichung = abweichung + abs(k - durchschnitt)
        abweichung = round(abweichung / float(len(value)), 2)

        # if difference is greater than durchschnitt + abweichung * 2 for 3 days in a row, we concider it as an anomaly
        # why 3 days in a row? because otherwise a single increase would an error. this would raise the error counter to over 200
        # why *2? because we need some buffer to use more water. like when it is summer or something like that
        if(difference > (durchschnitt + abweichung)*2 and value[len(value) - 2] > (durchschnitt + abweichung)*2 and value[len(value) - 3] > (durchschnitt + abweichung)*2):
            x = [i, round(durchschnitt + abweichung, 2), difference]
            printErrors.append(x)
            count = count + 1
    # this if statement is only to print the relevant information (only columns with errors)
    if not printErrors:
        continue
    print("Spalte: " + str(device) + ", Errors: " + str(len(printErrors)) + ", List: " + str(printErrors))
print("All errors: ", count)

# ---------------------- RULE CHAIN -----------------------
# Daten bereingen. die if mit WM brauchen wir nicht, genauso wie die Fehlerwerte mit über 3000. aber alle NA müssen rausgefiltert werden..
# Wichtig ist das die Datenbank folgende infos beihaltet. Aktueller wert beziehungsweise alle werte. ein Wert indem wir den Durchschnitt speichern können. Timestamp ist auch wichtig.#
# Value kommt rein. wird in die Datenbank hinzugefügt. Datenbank kalkuiert aufs neue den Durchschnitt etc. danach wird überprüft und gegebenenfalls eine Notification gesendet
# Es muss difference UND der gesamt Wert in der Datenbank gespeichert werden. Um bessere Sichtbarkeit zu gewährleisten.

# maschine learning?
# lineare regression?
