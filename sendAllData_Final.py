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
te = ["fGYaRNnZYnGB4aMpLjgf", "EjsO9hLaEYP2GzdKGnZJ", "0dqIjcqaBDHxGmV5paak", "UNTOS1HuO2O1FqTFj70o"]
te2 = ["DPB5XYM4Xuq072E1eerW", "FYzeg4LGnjj6ZkSw1nxU", "2DUwU6PkRclJ2i0hlkIU", "K3QHurhlGslw7oLgVQMA", "CBm3KWpGI3HSvAYMVjKk", "Ts7pt3LGxTj71n24L3Ey", "jgRWTq9CY9I41XmxUHGH", "AH0bNgc7giLz2M54iu7N", "B311ID2HRFd7sFJ8Evgo", "isKvxLTTU4mWMobJr8zO"]
te3 = ["kDLjlTc0R3K2nuYBh95c", "cDfJwbrEzPpdW24myBHd", "hnlpHeoIWREl2MfH7Fvt", "NwgyHaCuB95G7Zep3RG5", "qRhCNogqQdKqOk5AMWZ5", "YmmBdm6lr1VZn5NAoczJ", "A2ZFgQ1q9d3yS6R5bjAW", "q3DeM07bQCoOU3nrS0km", "oQwwxw7YLwyGSHXwVc5n", "HPZj783P3vKYk9ThYulj", "LaVnFRaRyHnCQwNlfpdM"]
te4 = [""]
te5 = ["ZyZ4s7t1OPzzzDcres6J", "gO2b1MNfzwF6yQEUljjT", "UDicGMp6jubZ8nNZjizy", "IZalECgJL8Fa4pEX7JbY"]
# Ascii -> A bis Z ist 65 bis 90. anfangen tun wir mit B, also 66
table = 67
table2 = 64
bol = 0
# ----------------------------------------------------------------- #
broker = "77.237.53.201"  # host name
port = 13883  # data listening port
timestampp = firstWorksheet["A"]

for i in range(maxLines):
    table = 75
    table2 = 66
    bol = 1
    if (i == 0):
        continue

    for device in te3:

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
            tableName = firstWorksheet["" + chr(table2) + chr(table)]
            ACCESS_TOKEN = device
            client1 = paho.Client("control1")  # create client object
            client1.on_publish = on_publish  # assign function to callback
            client1.username_pw_set(ACCESS_TOKEN)  # access token from thingsboard device
            client1.connect(broker, port)  # establish connection
            payload = {
                "ts": timestampp[i].value,
                "values": {
                    "value": tableName[i].value
                }
            }
            MQTT_MSG = json.dumps(payload)
            ret = client1.publish("v1/devices/me/telemetry", MQTT_MSG)  # topic-v1/devices/me/telemetry
            #print("Please check LATEST TELEMETRY field of your device")
            time.sleep(0.75)
            client1.disconnect()
        table = table + 1
    print(i)