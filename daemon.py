from openpyxl import load_workbook
import datetime

class mevent:
	def __init__(self, date,event_type,lid):
		self.date = date
		self.event_type = event_type
		self.lid = lid
	def show(self):
		output = ''
		output = '{}{}'.format(output,self.event_type)
		output = '{} of {}'.format(output,self.lid)		
		output = '{} at {}'.format(output,self.date.strftime('%m/%d/%y'))
		return output

wb = load_workbook('MouseInformation_AllLines_copy.xlsx')

def getCol(s,t):
	for cell in s[1]:
		if cell.value is not None:
			if cell.value.startswith(t):
				return cell.column
	return 0

def eventsfcol(des,lid,dates):
	mevents = []
	for cell in dates:
			if cell.value is not None:
				if type(cell.value) is type(datetime.datetime.now()):
					id = lid[cell.row-1]
					e = mevent(cell.value,des,id.value)
					mevents.append(e)
	return mevents				



local = []

for ws in wb:
	if getCol(ws,'Dissection Date') is not 0 and getCol(ws, 'Litter ID') is not 0:
		disevents = eventsfcol('Dissect',ws[getCol(ws,'Litter ID')], ws[getCol(ws,'Dissection Date')])
		for event in disevents:
			local.append(event)

	if getCol(ws,'Wean') is not 0 and getCol(ws, 'Litter ID') is not 0:
		winevents = eventsfcol('Wean',ws[getCol(ws,'Litter ID')], ws[getCol(ws,'Wean')])
		for event in winevents:
			local.append(event)

	if getCol(ws,'Tattoo') is not 0 and getCol(ws, 'Litter ID') is not 0:
		tatevents = eventsfcol('Tattoo',ws[getCol(ws,'Litter ID')], ws[getCol(ws,'Tattoo')])
		for event in tatevents:
			local.append(event)
		
for event in local:
	print event.show()
	print ''		
	

print 'check'





"""from apiclient.discovery import build
from httplib2 import Http
from oath2client import file, client, tools





from __future__ import print_function

try:
	import argparseu
	flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
	flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET = 'client_secret.json'

store = file.Storage('storage.json')
credz = store.get()

if not credz or credz.invalid:
	flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
	credz = tools.run_flow(flow, store, flags) \ 
		if flags else tools.run(flow,store)

API_KEY = 'AIzaSyCTswttl1PbzHtbD5oy3v7Cj1iILgAiI8k'

CAL = build('calendar', 'v3', http=credz.authorize(Http()))

GMT_OFF = '-04:00'

EVENT = {
	'summary': 'Dissect mouse',
	'start': {'dateTime': '2018-08-
}

"""
