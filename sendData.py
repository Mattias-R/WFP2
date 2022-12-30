import paho.mqtt.client as paho  # mqtt library
import os
import json
import time
from datetime import datetime

ACCESS_TOKEN = 'c6TzfPdjXR9bQCOfvAjk'  # Token of your device
broker = "77.237.53.201"  # host name
port = 13883  # data listening port


def on_publish(client, userdata, result):  # create function for callback
    print("data published to thingsboard \n")
    pass


client1 = paho.Client("control1")  # create client object
client1.on_publish = on_publish  # assign function to callback
client1.username_pw_set(ACCESS_TOKEN)  # access token from thingsboard device
#client1.username_pw_set("marbanriedl", "WFP1")
client1.connect(broker, port)  # establish connection

while True:
    payload = "{"
    #payload += "\"value\":60,";
    payload += "\"Temperature\":60";
    payload += "}"
    ret = client1.publish("v1/devices/me/telemetry", payload)  # topic-v1/devices/me/telemetry
    print("Please check LATEST TELEMETRY field of your device")
    print(payload);
    time.sleep(2)