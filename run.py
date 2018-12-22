from __future__ import print_function
from googleapiclient.discovery import build
from oauth2client import file, client, tools
from httplib2 import Http
from dateutil import parser

import csv
import json
import pandas as pd
from datetime import datetime

SCOPES = 'https://www.googleapis.com/auth/calendar'

FILE = "birthdays.csv"
BIRTHDAY_COL = "BIRTHDAY"
FIRSTNAME_COL = "FIRST NAME"
LASTNAME_COL = "LAST NAME"
STARTTIME_COL = "START TIME"
ENDTIME_COL = "END TIME"
YEAR=str(datetime.now().year)


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
        birthday_list = birth_date.split('/')
        birthday_list[-1] = YEAR
      
        
        birthday = "./".join(birthday_list)
        birthday = birthday.replace(' ','T')
        birthday = parser.parse(birthday)
        birthday = str(birthday).replace(' ', 'T')
        full_name = row[1][birthday_df.columns.get_loc(FIRSTNAME_COL)] +" "+ row[1][birthday_df.columns.get_loc(LASTNAME_COL)]
        
        info["name"] = full_name
        info["begin"] = json.dumps(birthday, default = convert_time)
        info["end"] = json.dumps(birthday.replace("00:00:00","23:59:00"), default = convert_time) 
        new = add_birthday(info)
        
        event = service.events().insert(calendarId='primary', body=new).execute()
        print('Birthday added: ', full_name)


def convert_time(time):
    if isinstance(time, datetime.datetime):
        return time.__str__()


def add_birthday(list):
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