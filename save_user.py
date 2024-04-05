#!/usr/bin/env python

import time
import RPi.GPIO as GPIO
from mfrc522 import SimpleMFRC522
import mysql.connector
import RPi_I2C_driver

db = mysql.connector.connect(
  host="localhost",
  user="attendanceadmin",
  passwd="f",
  database="attendancesystem"
)

cursor = db.cursor()
reader = SimpleMFRC522()
lcd = RPi_I2C_driver.lcd()
lcd.lcd_clear()

try:
  while True:
    lcd.lcd_clear()
    lcd.lcd_display_string('Place Card to',1)
    lcd.lcd_display_string('register',2)
    id, text = reader.read()
    cursor.execute("SELECT id FROM users WHERE rfid_uid="+str(id))
    cursor.fetchone()

    if cursor.rowcount >= 1:
      lcd.lcd_clear()
      lcd.lcd_display_string('Overwrite',1)
      lcd.lcd_display_string('existing user?',2)
      overwrite = input("Overwite (Y/N)? ")
      if overwrite[0] == 'Y' or overwrite[0] == 'y':
        lcd.lcd_clear()
        lcd.lcd_display_string("Overwriting user.",1)
        time.sleep(1)
        sql_insert = "UPDATE users SET name = %s WHERE rfid_uid=%s"
      else:
        continue;
    else:
      sql_insert = "INSERT INTO users (name, rfid_uid, cohort) VALUES (%s, %s, %s)"
    lcd.lcd_clear()
    lcd.lcd_display_string('Enter new name',1)
    new_name = input("Name: ")
    c_id = input("Cohort: ")
    cursor.execute(sql_insert, (new_name, id, c_id))

    db.commit()

    lcd.lcd_clear()
    lcd.lcd_display_string("User " + new_name,1)
    lcd.lcd_display_string('Saved',2)
    time.sleep(2)
finally:
  GPIO.cleanup()
