#Spreadsheet needs to be shared with 'clieant_email from oAuth .json file

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import datetime
import Tkinter as tk
from Tkinter import IntVar

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\Users\G\Desktop\pytime\GspreadTimeClock-e0468fff419c.json', scope)
gc = gspread.authorize(credentials)
sheet = gc.open('Time Sheet')
worksheet = sheet.get_worksheet(0)

def WhosPunchingCard(self):
	self = intvar.get()
	if self == 1:
		print 'Mitch'
		return 'Mitch'
	elif self == 2:
		print 'Izaac'
		return 'Izaac'
	elif self == 3:
		print 'Dan'
		return 'Dan '

def center(toplevel): 
    toplevel.update_idletasks()
    w = toplevel.winfo_screenwidth()
    h = toplevel.winfo_screenheight() 
    size = tuple(int(_) for _ in toplevel.geometry().split('+')[0].split('x'))
    x = w/2 - size[0]/2
    y = h/2 - size[1]/2
    toplevel.geometry("%dx%d+%d+%d" % (size + (x, y)))

def getRowColumnTuple(self):
	self = str(self)
	findRegex = str(worksheet.find(self))
	stringy = "'"
	stripSyntax = findRegex.rstrip('%s%s%s>' % (stringy, self, stringy))
	stripMoreSyntax = stripSyntax.lstrip('<Cell ')
	
	rowRegex = re.compile(r'R(\d{2}|\d)') #'(\D\d\d){2}|(\D\d){2}'
	colRegex = re.compile(r'C(\d{2}|\d)')
	searchRow = rowRegex.search(stripMoreSyntax)
	searchCol = colRegex.search(stripMoreSyntax)
	finalRowLocation = searchRow.group()
	finalColumnLocation = searchCol.group()

	justTheRowInt = str(finalRowLocation).lstrip('R')
	justTheColInt = str(finalColumnLocation).lstrip('C')
	
	return int(justTheRowInt), int(justTheColInt)

def getRowOrColumnWithTodaysDate():
	todayToTuple = getRowColumnTuple(str(datetime.date.today().strftime('X%m/X%d/20%y').replace('X0', '').replace('X', '')))
	return todayToTuple

def militaryTimestampWithNoSeconds():
	return datetime.datetime.now().strftime('%H:%M')

def getClockInCellAndEdit():
	employeesColumn = getRowColumnTuple(WhosPunchingCard(intvar.get()))[1]
	worksheet.update_cell(getRowOrColumnWithTodaysDate()[0], employeesColumn, militaryTimestampWithNoSeconds())

def getClockOutCellAndEdit():
	employeesColumn = int(getRowColumnTuple(WhosPunchingCard(intvar.get()))[1]) + 1
	worksheet.update_cell(getRowOrColumnWithTodaysDate()[0], employeesColumn, militaryTimestampWithNoSeconds())

class MainApp(tk.Tk):
	def __init__(self):
		tk.Tk.__init__(self)
		global intvar
		intvar = IntVar()

		ClockInButton = tk.Button(width=30, height=5, text='ClockIn', command=lambda: getClockInCellAndEdit())
		ClockInButton.grid(row=1, column=1, padx=(20,5), pady=(25,25))

		ClockOutButton = tk.Button(width=30, height=5, text='ClockOut', command=lambda: getClockOutCellAndEdit())
		ClockOutButton.grid(row=1, column=3, padx=(5,20), pady=(25,25))

		MitchellButton = tk.Radiobutton(text="User1", width=5, value=1, variable=intvar)
		MitchellButton.grid(row=0, column=1, padx=1, pady=25)

		IzaacButton = tk.Radiobutton(text="User2", width=5, value=2, variable=intvar)
		IzaacButton.grid(row=0, column=2, padx=1, pady=25)

		DanButton = tk.Radiobutton(text="User2", width=5, value=3, variable=intvar)
		DanButton.grid(row=0, column=3, padx=1, pady=25)

if __name__ == "__main__":
	root = MainApp()
	#root.resizable(0,0)
	center(root)
	root.title('TimeClok App')
	root.mainloop()