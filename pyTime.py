from oauth2client.service_account import ServiceAccountCredentials
import datetime, gspread, re
import Tkinter as tk
from Tkinter import IntVar

def center(toplevel): # Center root Tkinter GUI Window, respective to monitor resolution.
    toplevel.update_idletasks()
    w = toplevel.winfo_screenwidth()
    h = toplevel.winfo_screenheight() 
    size = tuple(int(_) for _ in toplevel.geometry().split('+')[0].split('x'))
    x = w/2 - size[0]/2
    y = h/2 - size[1]/2
    toplevel.geometry("%dx%d+%d+%d" % (size + (x, y)))

scope = ['https://spreadsheets.google.com/feeds'] #line 17-21 are GoogleSheets API functions for edtting permissions on specific sheet names.
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\Users\G\Desktop\pytime\GspreadTimeClock-e0468fff419c.json', scope)
gc = gspread.authorize(credentials)
sheet = gc.open('Time Sheet')
worksheet = sheet.get_worksheet(0) # process, work on the first tab within above spreadsheet.

names = ['Mitch','Izaac','Dan ']

def WhosPunchingCard(intIn): # pass integer variable from selected radio button
	if intIn == 1:
		return names[0] # depending on integer, pass name of employee from names array
	elif intIn == 2:
		return names[1] # could easily be employee ID number
	elif intIn == 3:
		return names[2]

def getRowColumnTuple(self): # ugly code for stripping/regex/parsing hardcoded info from the google sheets.
	self = str(self) # 
	findRegex = str(worksheet.find(self)) 
	stringy = "'" #google response fluff
	stripSyntax = findRegex.rstrip('%s%s%s>' % (stringy, self, stringy)) #strip excess characters from Google's response.
	stripMoreSyntax = stripSyntax.lstrip('<Cell ') # strip Google response, to bare minimum info.
	
	rowRegex = re.compile(r'R(\d{2}|\d)') # regex parameters for parsing googles API internal response. 
	colRegex = re.compile(r'C(\d{2}|\d)')
	searchRow = rowRegex.search(stripMoreSyntax) # actually search, using regex as cross reference.
	searchCol = colRegex.search(stripMoreSyntax)
	finalRowLocation = searchRow.group() # string output of regex search, when matched.
	finalColumnLocation = searchCol.group()

	justTheRowInt = str(finalRowLocation).lstrip('R') # take off R/C string character, leave value. 
	justTheColInt = str(finalColumnLocation).lstrip('C')
	
	return int(justTheRowInt), int(justTheColInt) # return row, column tuple of found/matched cells corresponding to employee name, and todays date on the spreadsheet.

def getRowOrColumnWithTodaysDate(): #pass specially formatted datetime to row/column function. Which gets coordinates of cell corresponding to todays date.
	todayToTuple = getRowColumnTuple(str(datetime.date.today().strftime('X%m/X%d/20%y').replace('X0', '').replace('X', '')))
	return todayToTuple

def militaryTimestampWithNoSeconds(): # return current time. For clock in/out stamp.
	return datetime.datetime.now().strftime('%H:%M')

def getClockInCellAndEdit(intIn): # the final goods! line 61 is google API command being passed, coordinates of cell with todays date, employee column and current time.
	employeesColumn = getRowColumnTuple(WhosPunchingCard(intIn))[1] # get row/column coordinates of employee who is punching time card.
	worksheet.update_cell(getRowOrColumnWithTodaysDate()[0], employeesColumn, militaryTimestampWithNoSeconds()) # passing column of cell, row of cell, and what to update it with.
	print(str(names[intIn-1])) #debug, sanity check.

def getClockOutCellAndEdit(intIn): # Final goods, for clock out.
	employeesColumn = int(getRowColumnTuple(WhosPunchingCard(intIn))[1]) + 1 # Same as clock in but plus one. eg. Clock in is on column 5, clock out is column 6. 
	worksheet.update_cell(getRowOrColumnWithTodaysDate()[0], employeesColumn, militaryTimestampWithNoSeconds())
	print(str(names[intIn-1])) #debug, sanity check.

class MainApp(tk.Tk):
	def __init__(self):
		tk.Tk.__init__(self)
		intvar = IntVar() # assign new integer variable (+1) with every usage.
		
		MitchellButton = tk.Radiobutton(text="Mitchell", width=5, value=1, variable=intvar)
		MitchellButton.grid(row=0, column=1, padx=1, pady=25)
		#MitchellButton.configure(background='grey')
		#MitchellButton.bind("<1>", intvar.get())

		IzaacButton = tk.Radiobutton(text="Izaac", width=5, value=2, variable=intvar)
		IzaacButton.grid(row=0, column=2, padx=1, pady=25)
		#IzaacButton.configure(background='grey')
		#IzaacButton.bind("<2>", intvar.get())

		DanButton = tk.Radiobutton(text="Dan", width=5, value=3, variable=intvar)
		DanButton.grid(row=0, column=3, padx=1, pady=25)
		#DanButton.configure(background='grey')
		#DanButton.bind("<3>", intvar.get())

		stringInput = tk.Entry(width=85, relief='sunken')#, command=lambda: stringEntryToggleInOut())
		stringInput.grid(row=1, column=1, columnspan=3, padx=10, pady=5)
		stringInput.focus()

		ClockInButton = tk.Button(width=30, height=5, text='ClockIn', command=lambda: getClockInCellAndEdit(intvar.get())) # pass radiobutton integer to Clock in function.
		ClockInButton.grid(row=2, column=1, padx=(20,5), pady=(25,25))
		#ClockInButton.bind("<+>", getClockInCellAndEdit)

		ClockOutButton = tk.Button(width=30, height=5, text='ClockOut', command=lambda: getClockOutCellAndEdit(intvar.get())) # pass radiobutton integer to clock out function.
		ClockOutButton.grid(row=2, column=3, padx=(5,20), pady=(25,25))
		#ClockOutButton.bind("<->", getClockInCellAndEdit)

if __name__ == "__main__":
	root = MainApp()
	root.resizable(0,0)
	center(root)
	root.title('TimeClok App')
	#root.configure(background='grey')
	root.mainloop()
