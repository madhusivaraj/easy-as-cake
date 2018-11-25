from __future__ import print_function
from googleapiclient.discovery import build
from oauth2client import file, client, tools
from httplib2 import Http
from dateutil import parser

import csv
import json
import pandas as pd
import datetime

SCOPES = 'https://www.googleapis.com/auth/calendar'

FILE = "birthdays.csv"
BIRTHDAY_COL = "BIRTHDAY"
FIRSTNAME_COL = "FIRST NAME"
LASTNAME_COL = "LAST NAME"
STARTTIME_COL = "START TIME"
ENDTIME_COL = "END TIME"
YEAR='2018'

def main():
    birthday_df = pd.read_csv(FILE)
    store = file.Storage('token.json')
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        credentials = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=credentials.authorize(Http()))
    
    for row in birthday_df.iterrows():
        info = {} 
        birth_date = row[1][birthday_df.columns.get_loc(BIRTHDAY_COL)]
        birthdayList = birth_date.split('/')
        
        if(int(birthdayList[0])>11):
            birthdayList[-1] = YEAR
        else:
            birthdayList[-1] = '2019'
        
        birthday = "./".join(birthdayList)
        birthday = birthday.replace(' ','T')
        birthday = parser.parse(birthday)
        birthday = str(birthday).replace(' ', 'T')
        fullName = row[1][birthday_df.columns.get_loc(FIRSTNAME_COL)] +" "+ row[1][birthday_df.columns.get_loc(LASTNAME_COL)]
        
        info["name"] = fullName
        info["begin"] = json.dumps(birthday, default = timeConversion)
        info["end"] = json.dumps(birthday.replace("00:00:00","23:59:00"), default = timeConversion) 
        new = addBirthday(info)
        
        event = service.events().insert(calendarId='primary', body=new).execute()
        print('Event added: ', fullName)

def timeConversion(time):
    if isinstance(time, datetime.datetime):
        return time.__str__()

def addBirthday(list):
    TIMEZONE = 'America/New_York'
    RRule = 'RRULE:FREQ=YEARLY'
    event = {
      'summary': list["name"] + '\'s Birthday',
      'start': {
        'dateTime': list["begin"].replace("\"",""),
        'timeZone': TIMEZONE,
      },
      'end': {
        'dateTime': list["end"].replace("\"",""),
        'timeZone': TIMEZONE,
      },
      'reminders': {
        'useDefault': True,
      },
      'recurrence': [RRule],
      'colorId': '7'
    }
    return event


if __name__ == '__main__':
    main()