#!/usr/bin/env python
# coding: utf-8

# In[1]:


from __future__ import print_function

import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


import pandas as pd
from IPython.display import display
import json


# In[2]:


"""
Change pandas dataframe option.
"""

pd.set_option('display.max_colwidth', None)


# In[3]:


"""
Help functions:
- Convert json string to DataFrame.
- Convert Dict to DataFrame.
- Display DataFrame as a Table using IPython library.

"""


def jsonStringToDataFrame(jsonString):
    """
    Convert jsonString to dict.
    """
    jsonStr = str(jsonString)
    strAsJson = json.loads(jsonStr)
    
    # Use pandas.DataFrame.from_dict() to Convert JSON to DataFrame
    dictObj = pd.DataFrame.from_dict(strAsJson, orient="index")
    return dictObj

def dictToDataFrame(Dict):
     # Use pandas.DataFrame.from_dict() to Convert dict to DataFrame
    dictObj = pd.DataFrame.from_dict(Dict, orient="index")
    return dictObj

def displaydf(df):
    """
    Display df as a table
    """
    display(df)


# In[4]:


"""
Set the API Scoopes
"""
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly',
          'https://www.googleapis.com/auth/calendar.events',
          'https://www.googleapis.com/auth/calendar']


# In[5]:


"""
#Get_events from user calendar.

Shows basic usage of the Google Calendar API.
Prints the start and name of the next 10 events on the user's calendar.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

# displaydf(jsonStringToDataFrame(creds.to_json()))


# In[6]:


print(creds.expiry)
print(creds.granted_scopes)


# In[7]:


"""
If his is the first time the user uses this service, 
let him login and ask for needed permissions,
then save his token. 
"""
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())


# In[8]:


"""
The different between printing the base creds and the converted one. 
"""
print("Before:\n")
print(creds.to_json())
print("\n\nAfter:\n")
displaydf(jsonStringToDataFrame(creds.to_json()))


# In[9]:


"""
Get the 
"""

try:
    service = build('calendar', 'v3', credentials=creds)# set the api parameter

    # Call the Calendar API 
    now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    print('Getting the upcoming 10 events:\n')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=10, singleEvents=True,
                                          orderBy='startTime').execute()#get the next 10 events.
   
    events = events_result.get('items', [])# save the events in a list.

    if not events:
        print('No upcoming events found.')
        

    # Prints the start and name of the next 10 events, and write it to an excel sheet.
    startRow=0
    destination = 'eventsLog.xlsx'
    writer = pd.ExcelWriter(destination, engine='openpyxl')
    
    for event in events:
        print(type(event))
        eventDF = dictToDataFrame(event)
        start = event['start'].get('dateTime', event['start'].get('date'))
        ev = event['summary']
        print(f'Event {ev}')
        sheetname = event['summary']
        eventDF.to_excel(writer, sheet_name= sheetname[:5])
        startRow+=1
        displaydf(eventDF)
        print(start,' ', event['summary'],'\n')
    writer.save()   
except HttpError as error:
    print('An error occurred: %s' % error)

# Refer to the Python quickstart on how to setup the environment:
# https://developers.google.com/calendar/quickstart/python
# Change the scope to 'https://www.googleapis.com/auth/calendar' and delete any
# stored credentials.


# In[10]:


"""
Create event json and isert it to the user calendar.
"""

event = {
'summary': 'Google I/O 2022',
'location': '800 Howard St., San Francisco, CA 94103',
'description': 'A chance to hear more about Google\'s developer products.',
'start': {
'dateTime': '2022-12-10T09:00:00-07:00',
'timeZone': 'America/Los_Angeles',
},
'end': {
'dateTime': '2022-12-11T09:00:00-08:00',
'timeZone': 'America/Los_Angeles',
},
'recurrence': [
'RRULE:FREQ=DAILY;COUNT=2'
],
'attendees': [
{'email': 'lpage@example.com'},
{'email': 'sbrin@example.com'},
],
'reminders': {
'useDefault': False,
'overrides': [
    {'method': 'email', 'minutes': 24 * 60},
    {'method': 'popup', 'minutes': 10},
],
},
}

eventsList = []
for i in range(5): 
    event = {
            f'summary': 'Testing calendar API No:' + str(i),
            'location': '800 Howard St., San Francisco, CA 94103',
            'description': 'A chance to hear more about Google\'s developer products.',
            'start': {
            'dateTime': '2022-12-10T09:00:00-07:00',
            'timeZone': 'America/Los_Angeles',
            },
            'end': {
            'dateTime': '2022-12-11T09:00:00-08:00',
            'timeZone': 'America/Los_Angeles',
            },
            'recurrence': [
            'RRULE:FREQ=DAILY;COUNT=2'
            ],
            'attendees': [
            {'email': 'lpage@example.com'},
            {'email': 'sbrin@example.com'},
            ],
            'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
            },
        }
    eventsList.append(event)

for event in eventsList:
    
    event = service.events().insert(calendarId='primary', body=event).execute()
    print ('Event created: %s' % (event.get('htmlLink')))


# In[ ]:





# In[11]:


"""
in this cell we will query the free time in the user calendar.
"""
the_datetime = '2022-12-20T15:00:00Z'
the_datetime2 = '2022-12-22T16:00:00Z'
calendarId = 'mohammed.s.g76@gmail.com'

#DateTime should be on the RFC33393 standard
def calendarFreeAt(fromDateTime:str, toDateTime:str, calendarId:str):
    """
    This function should return True if the selected time is free.
    """
    body = {
      "timeMin": fromDateTime,
      "timeMax": toDateTime,
      "timeZone": 'UTC',
      "items": [{"id": calendarId}]
    }


    eventsResult = service.freebusy().query(body=body).execute()
    
    displaydf(dictToDataFrame(eventsResult))

    cal_dict = eventsResult[u'calendars']
    if len(cal_dict[calendarId]['busy']) == 0:
        print('this time is free of events')
        return True
    else:
        for cal_name in cal_dict:
            displaydf(dictToDataFrame(cal_dict))
            print(cal_name, cal_dict[cal_name])
        return False
    
calendarFreeAt(the_datetime, the_datetime2, calendarId)


# In[12]:


the_datetime = '2022-12-20T15:00:00Z'
the_datetime2 = '2022-12-22T16:00:00Z'

body = {
  "timeMin": the_datetime,
  "timeMax": the_datetime2,
  "timeZone": 'UTC',
  "items": [{"id": 'mohammed.s.g76@gmail.com'}]
}


eventsResult = service.freebusy().query(body=body).execute()

# displaydf(dictToDataFrame(eventsResult))

cal_dict = eventsResult[u'calendars']
# displaydf(dictToDataFrame(cal_dict))
for cal_name in cal_dict:
    print(cal_name, cal_dict[cal_name]['busy'])
    eventsList = cal_dict[cal_name]['busy']
    print(type(eventsList))
    if len(eventsList) == 0:
        print('free')
event_dict = cal_dict.get('mohammed.s.g76@gmail.com')

print(event_dict.get('busy'))


# In[ ]:




