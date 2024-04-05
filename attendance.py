import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


import RPi.GPIO as GPIO
from mfrc522 import SimpleMFRC522
import mysql.connector
import RPi_I2C_driver
from subprocess import check_output

import pandas as pd
import numpy as np
import string
import time
from datetime import datetime
import sys
import math
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "ENTER SPREADHSEET ID HERE"

db = mysql.connector.connect(
    host="localhost",
    user="attendanceadmin",
    passwd="f",
    database="attendancesystem"
)

creds = None
SHEET = "SHEET NAME"

cursor = db.cursor()
reader = SimpleMFRC522()

buzzer = 16
GPIO.setup(buzzer, GPIO.OUT)
p = GPIO.PWM(buzzer,1)
p.start(0)

lcd = RPi_I2C_driver.lcd()

# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "cred.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())
try:
    service = build("sheets", "v4", credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SHEET+"!$A$1:YY")
        .execute()
    )
    values = result.get("values", [])
    global df
    df = pd.DataFrame(values)
except HttpError as err:
    print(err)



prevtime = time.time()
prevuser = ""
global prevloc1,prevloc2
prevloc1 = (1,1)
prevloc2 = (1,1)

def playSound(tone, t):
    p.start(50)
    sinVal = tone
    toneVal = 2000 + sinVal*500
    p.ChangeFrequency(toneVal)
    time.sleep(t)
    p.stop()

def convert_coordinates(sheet_name,start_coord = (1,1), end_coord = (1,1)):
    # Convert numerical column index to letter
    def get_column_letter(col_num):
        div = col_num
        column_letter = ""
        while div > 0:
            (div, mod) = divmod(div - 1, 26)  # 26 letters in the alphabet
            column_letter = string.ascii_uppercase[mod] + column_letter
        return column_letter


    # Extract row and column indices from coordinates
    start_row, start_col = start_coord
    end_row, end_col = end_coord


    # Convert column indices to letters
    start_col_letter = get_column_letter(start_col)
    end_col_letter = get_column_letter(end_col)


    # Construct the range string
    range_string = f""
    if (start_coord == end_coord):
        range_string = f"{sheet_name}!{start_col_letter}{start_row}"
    else:
        range_string = f"{sheet_name}!{start_col_letter}{start_row}:{end_col_letter}{end_row}"


    return range_string


def get_sheets():
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SHEET+"!$A$1:YY")
        .execute()
    )
    values = result.get("values", [])
    global df
    df = pd.DataFrame(values)

def save_sheets(student_name):
    df1 = pd.DataFrame(df)
    df1.replace(to_replace = np.nan, value = '', inplace=True)
    # print (prevloc1)
    if (prevloc1 == prevloc2):
        result["values"] = [[str(datetime.now())]]
    else:
        result["values"] = [[student_name, str(datetime.now())]]
    # print (result["values"])
    rangestr = convert_coordinates(SHEET, prevloc1, prevloc2)
    # rangestr1 = convert_coordinates(SHEET, (1,1), (1,1))
    # print(rangestr)
    result["range"] = rangestr
    sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=rangestr, valueInputOption = "USER_ENTERED", body=result).execute()
  
def input_attendance(student_name):
    get_sheets()
    global prevloc1, prevloc2
    if student_name in df[0].values:
        timestamp = str(datetime.now())
        cell1 = df.loc[df[0] == student_name].iloc[0].iloc[len(df.axes[1])-1]
        if ((cell1 == '') or (type(cell1) != str)):
            for i in range(1, len(df.axes[1])):
                currval = df[df[0] == student_name].iloc[0].iloc[i]
                if currval == '' or type(currval) != str:
                    # df.loc[df[0] == student_name, i] = timestamp
                    prevloc1 = (int(df[df[0] == student_name].index[0])+1, i+1)
                    prevloc2 = prevloc1
                    break
        else:
            # df.loc[df[0] == student_name, len(df.axes[1])] = timestamp
            prevloc1 = (int(df[df[0] == student_name].index[0])+1, int(len(df.axes[1]))+1)
            prevloc2 = prevloc1
    else:
        # newrow = [student_name, str(datetime.now())]
        prevloc1 = (len(df)+1, 1)
        prevloc2 = (len(df)+1, 2)
        # df.loc[len(df)] = pd.Series(newrow, index=df.columns[:len(newrow)])
    save_sheets(student_name)

def main():
    ipaddr = check_output(['hostname','-I']).decode().strip()
    lcd.lcd_clear()
    lcd.lcd_display_string(str(ipaddr),1)
    time.sleep(2)
    playSound(3,1)
    global SHEET
    try:
        global prevtime
        global prevuser
        while True:
            lcd.lcd_clear()
            lcd.lcd_display_string('Tap to record',1)
            lcd.lcd_display_string('attendance', 2)
            id, text = reader.read()

            cursor.execute("Select id, name, cohort FROM users WHERE rfid_uid="+str(id))
            result1 = cursor.fetchone()

            lcd.lcd_clear()
            if cursor.rowcount >= 1:
                if (prevuser == result1[1]) and (time.time() - prevtime < 20): # dont record same user twice
                    lcd.lcd_display_string("Attendance",1)
                    lcd.lcd_display_string("Already Recorded",2)
                    playSound(0.7,0.15)
                    time.sleep(0.07)
                    playSound(0.7,0.15)
                else:
                    lcd.lcd_display_string("Welcome",1)
                    lcd.lcd_display_string(result1[1], 2)
                    cursor.execute("INSERT INTO attendance (user_id) VALUES (%s)", (result1[0],) )
                    if (result1[2] == 2):
                        SHEET = "Cohort2"
                    else:
                        SHEET = "Cohort1"
                    input_attendance(result1[1])
                    db.commit()
                    prevtime = time.time()
                    prevuser = result1[1]
                    playSound(3,0.15)
                    time.sleep(0.07)
                    playSound(7.2,0.15)
            else:
                lcd.lcd_display_string('User does not',1)
                lcd.lcd_display_string('exist', 2)
                playSound(2.85,0.15)
                time.sleep(0.07)
                playSound(0.7,0.15)
            time.sleep(1.6)
    except KeyboardInterrupt:
        lcd.lcd_clear()
        GPIO.cleanup()
        sys.exit()
    except:
        print("Error has occured")
        lcd.lcd_clear()
        lcd.lcd_display_string('Error! Please',1)
        lcd.lcd_display_string('request help.', 2)
        for a in range(6):
            playSound(2.85,0.1)
            time.sleep(0.05)
        time.sleep(3)
        main()
    finally:
        lcd.lcd_clear()
        GPIO.cleanup()

if __name__ == "__main__":
   main()

