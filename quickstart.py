from __future__ import print_function
import datetime
import xlrd
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_id.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))


    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    start_hh=['08','09','11','14','15','17']
    start_mm=['00','50','40','00','45','30']
    end_hh=['09','11','13','15','17','19']
    end_mm=['35','25','15','35','20','05']
    book=xlrd.open_workbook('imi2018.xls')
    it4=book.sheet_by_index(8)
    for i in range(3,39):
        if it4.cell(i, 8).value == "":
            continue                                              
        event = {
          'summary': it4.cell(i, 8).value,
          'location': it4.cell(i, 10).value,
          'description': it4.cell(i, 9).value,
          'start': {
            'dateTime': f"2018-10-08T{start_hh}:{start_mm}:00+09:00",
            'timeZone': 'Asia/Yakutsk',
          },
          'end': {
            'dateTime': f"2018-10-08T{end_hh}:{end_mm}:00+09:00",
            'timeZone': 'Asia/Yakutsk',
          },
          'recurrence': [
            'RRULE:FREQ=WEEKLY;COUNT=12'
          ],
          'reminders': {
          }
        }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
    main()