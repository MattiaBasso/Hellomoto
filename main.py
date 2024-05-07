import datetime
import numpy as np
import win32com.client as win32
import pandas as pd
import openpyxl
outlook = win32.gencache.EnsureDispatch("Outlook.Application")

start = "2024-05-07 16:10" # yyyy-MM-dd hh:mm


def send_meeting(scadenza, oggetto):
    event = outlook.CreateItem(1) # AppointmentItem
    event.Start = scadenza
    event.Subject = oggetto
    event.Duration = 60  # In minutes (60 Minutes)
    event.Location = ''
    event.Body = ''
    # event.AllDayEvent = True
    event.Save()
    event.Send()


df = pd.read_excel(r"C:\Users\mattia.basso\OneDrive - Basso s.r.l\Desktop\NUOVA DELHIZIA.xlsx", sheet_name="REGISTRO", parse_dates=["SCADENZA"], dtype={"FLAG_CALENDARIO": 'boolean'})
df["FLAG_CALENDARIO"] = df["FLAG_CALENDARIO"].fillna(value=False)
date = df[df["SCADENZA"] > datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
for row, data in date.iterrows():
    if not data["FLAG_CALENDARIO"] or data["FLAG_CALENDARIO"] == pd.NA:
        print("Done")
        # send_meeting(data.SCADENZA, data.LAVORO)
        data.FLAG_CALENDARIO = True
df.update(date)
print(date.FLAG_CALENDARIO.head())
