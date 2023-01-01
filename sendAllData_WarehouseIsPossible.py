# importing openpyxl module
import openpyxl
import paho.mqtt.client as paho  # mqtt library
import os
import json
import time
from datetime import datetime

def on_publish(client, userdata, result):  # create function for callback
    print("data published to thingsboard \n")
    pass

def send_data(device, timestampp, tableName):
    ACCESS_TOKEN = device
    client1 = paho.Client("control1")  # create client object
    client1.on_publish = on_publish  # assign function to callback
    client1.username_pw_set(ACCESS_TOKEN)  # access token from thingsboard device
    client1.connect(broker, port)  # establish connection
    payload = {
        "ts": timestampp,
        "values": {
            "value": tableName
        }
    }
    MQTT_MSG = json.dumps(payload)
    ret = client1.publish("v1/devices/me/telemetry", MQTT_MSG)  # topic-v1/devices/me/telemetry
    #print("Please check LATEST TELEMETRY field of your device")
    time.sleep(1)
    client1.disconnect()



# input excel file path
inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_Neuhaus.xlsx'
# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)
sheet_obj = newWorkbook.active
# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["Tabelle1"]
maxLines = sheet_obj.max_row

#Token for sending data
token = ["lIsqlNF34ZiGMphYoyUH", "ga9oIwWhnyot0PkZCAbZ", "F0XGyDOMlgB5OExB7lPa", "PXiWHn07oAZcyKjxpu4r", "6yzntB4m2qcfqHvfSE7i"]

# Ascii -> A bis Z ist 65 bis 90. anfangen tun wir mit B, also 66
table = 67
table2 = 64
bol = 0
# ----------------------------------------------------------------- #
broker = "77.237.53.201"  # host name
port = 13883  # data listening port
timestampp = firstWorksheet["A"]

for i in range(maxLines):
    table = 67
    table2 = 64
    bol = 0
    if (i == 0):
        continue

    for device in token:

        if (table == 91):
            table = 65
            bol = 1
            table2 = table2 + 1

        if (bol == 1):
            tableName = firstWorksheet["" + chr(table2) + chr(table)]
            print(chr(table2) + chr(table))
        else:
            tableName = firstWorksheet["" + chr(table)]

        if (tableName[i].value != "NA" and float(tableName[i].value) > 3000):
            tableName[i].value = "NA"

        # skip loop if value is NA
        if (tableName[i].value == "NA"):
            print("na")
        else:
            tableName = firstWorksheet["" + chr(table)]
            send_data(device,timestampp[i].value, tableName[i].value)
            print(device, timestampp[i].value, tableName[i].value)
        time.sleep(2)
        table = table + 1
    print(i)