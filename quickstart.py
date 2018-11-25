from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def main():
  
    store = file.Storage('token.json')
    creds = store.get()

    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    GCAL = build('calendar', 'v3', http=creds.authorize(Http()))

    # Call the Calendar API
    
    TIMEZONE = 'America/New_York'
    RRule = 'RRULE:FREQ=WEEKLY;COUNT=15'
    EVENT = {
      'summary': 'Dinner with friends',
      'start':   {'dateTime': '2018-09-15T19:00:00','timeZone': TIMEZONE},
      'end':     {'dateTime': '2018-09-15T22:00:00','timeZone': TIMEZONE},
      'recurrence': [RRule]
    }

    e = GCAL.events().insert(calendarId='primary',body=EVENT).execute()

    

    print('''*** %r event added:
    Start: %s
    End:   %s''' % (e['summary'].encode('utf-8'),
        e['start']['dateTime'], e['end']['dateTime']))

    print('URL: {}'.format(e['htmlLink']))


if __name__ == '__main__':
    main()