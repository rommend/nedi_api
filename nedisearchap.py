#!/usr/bin/env python

import requests
import sys
import xlsxwriter
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
print(sys.version)

hostInput = 'AIR-LAP1131AG-E-K9'

nediUrl = 'url'
nediUser = 'user'
nediPass = 'pass'

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

hostInput = "'%s'" % hostInput
nediTable = 'devices'
nediQuery = 'type'
nediOperatorStart = '='
nediOperatorEnd = '='
dataInput = 't=' + nediTable + '&q' + nediOperatorStart + nediQuery + nediOperatorEnd + hostInput
resp = requests.get(nediUrl + dataInput, auth=(nediUser, nediPass), verify=False).json()

for item in resp[1:]:
	#print(item['device'])
	#print(item['type'])
	row += 1
	worksheet.write(row, col, item['device'])
	worksheet.write(row, col + 1, item['type'])
	nediTable = 'links'
	nediQuery = 'device'
	device = item['device']
	device = "'%s'" % device
	dataInput = 't=' + nediTable + '&q' + nediOperatorStart + nediQuery + nediOperatorEnd + device
	resp = requests.get(nediUrl + dataInput, auth=(nediUser, nediPass), verify=False).json()
	try:
		#print(resp[1]['neighbor'])
		worksheet.write(row, col + 2, resp[1]['neighbor'])
		#print(resp[1]['nbrifname'])
		worksheet.write(row, col + 3, resp[1]['nbrifname'])
		neighbor = resp[1]['neighbor']
		hostInput = "'%s'" % neighbor
		nediTable = 'devices'
		nediQuery = 'device'
		dataInput = 't=' + nediTable + '&q' + nediOperatorStart + nediQuery + nediOperatorEnd + hostInput
		resp = requests.get(nediUrl + dataInput, auth=(nediUser, nediPass), verify=False).json()
		#print(resp[1]['type'])
		worksheet.write(row, col + 4, resp[1]['type'])
	except Exception as e:
		#print('no neighbor found')
		worksheet.write(row, col + 2, 'no neighbor found')

workbook.close()
exit()