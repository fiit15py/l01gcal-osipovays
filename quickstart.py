from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import xlrd

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def main():
	store = file.Storage('token.json')
	creds = store.get()
	if not creds or creds.invalid:
		flow = client.flow_from_clientsecrets('client_secret_524244190620-oipotnegi3fa7u6me41jig3vgnpv1pt8.apps.googleusercontent.com.json', SCOPES)
		creds = tools.run_flow(flow, store)
	service = build('calendar', 'v3', http=creds.authorize(Http()))
	# Call the Calendar API
	now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
	#############################################################################################
	start = ['08:00', '09:50', '11:40', '14:00', '15:45', '17:30']
	end = ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']  
	book = xlrd.open_workbook('imi2018.xls')                      
	it4 = book.sheet_by_index(8)                                  
																	  
	for i in range(3,39):                                         
		if it4.cell(i, 8).value == "":                            
			continue                                              
		para = it4.cell(i, 8).value                               
		l_pr = it4.cell(i, 9).value                               
		room = it4.cell(i, 10).value 
		#############################################################################################

		event = {
		  'summary': para,
		  'location': room,
		  'description': l_pr,
		  'start': {
			'dateTime': '2018-10-'+str((i-3)//6+1)+'T'+start[(i-3)%6]+':00+09:00',
			'timeZone': 'Asia/Yakutsk',
		  },
		  'end': {
			'dateTime': '2018-10-'+str((i-3)//6+1)+'T'+end[(i-3)%6]+':00+09:00',
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