# pyTimeClok
tKinter GUI for time stamping Google Sheets<br>
<br>
<img src=https://github.com/BiTinerary/pyTimeClok/blob/master/pyTimeImage.png>

## Over all
Searches hardcoded name of spreadsheet for a cell matching end user inputs. ie: Employee name<br>
* This parameter corresponds to respective users timecard/column.</br>

Does the same for locating today's date regex, which in turn provides row location.

Uses above row/column tuple as cell location for "Clock In and Clockout" buttons.</br>
* Default Tuple == Clock In</br>
* DefaultTuple[1] + 1 == ClockOut</br>
* These Tkinter buttons contain lambda functions that execute respective Clock In/Out Functions.</br>
    
###TODO
<strike>Add hardcoded "Employee Name" radiobuttons.</strike></br>
Then bind names/buttons to keyboard presses.</br>
    <tab>ie: <kbd>1</kbd> == User 1, <kbd>+</kbd> == Clock In, <kbd>-</kbd> == ClockOut</br>
