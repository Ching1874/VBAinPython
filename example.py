from Excel import Application
import threading
from pathlib import Path
'''
part 1 - preparation
'''
# instantiate the class Application
xlApp = Application(clsid="Excel.Application")

'''
part 2 - functions
remember to add xlApp.threading_event.set() at the end of each function
'''
# job 1: input value to the cell "A1"
def inputCellA1(inputText):
    xl = xlApp.interface
    xl.ActiveSheet.Range("A1").Value = inputText
    xlApp.threading_event.set()
# job 2: print the value of cell "A1" in python
def printCellA1():
    xl = xlApp.interface
    print(xl.ActiveSheet.Range("A1").Value)
    xlApp.threading_event.set()

'''
part 3 - kick the program to start
'''
excel_file_name = Path("example.xlsx").resolve()
try:
    xlApp.open(file=excel_file_name.as_posix()) # open excel
    xlApp.add_task(inputCellA1, "this is an example") # add step 1
    xlApp.add_task(printCellA1) # add step 2
    xlApp.start() # start the program
except Exception as e:
    print(f"- Error: {e}.")