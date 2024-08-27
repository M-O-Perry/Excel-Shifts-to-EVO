import datetime
from PlayActions import send_keys as send

from tkinter import filedialog, simpledialog, messagebox
import os
import openpyxl
from EVOUtil import openTASProgram

def getInputs():
    # get employee ID
    employeeID = simpledialog.askinteger("Employee ID", "Enter the employee ID")
    
    if employeeID is None:
        os._exit(0)
    
    file = filedialog.askopenfilename()
    if file == "":
        os._exit(0)
    
    return employeeID, file

employeeID, file = getInputs()

def getTimes(excelFile):
    wb = openpyxl.load_workbook(excelFile, data_only=True)
    sheet = wb.active

    times = []
    for i in range(3, 20):
        hours = sheet.cell(row = i, column = 24).value
        
        if hours is not None and is_float(hours) and float(hours) > 0:
            op = sheet.cell(row = i, column = 22).value
            sequence = sheet.cell(row = i, column = 23).value
            
            if op is not None and sequence is not None:      
                times.append([op, sequence, hours])
            
    wb.close()
    return times


# def openShifts():
#     openTASProgram("WOF")
#     send(["focus WO-M", 5])
#     send(["tab", "right 2", 0.5, "enter", 0.5])

# def addShift(employeeID, operation, sequence, hours, date = ""):
#     if date == "":
#         send(["enter"])
#     else:
#         send(["#"+str(date), "enter"])
    
#     send(["#"+str(employeeID), "enter", "#"+str(operation), "enter", "#"+str(sequence), "enter 4", "#"+str(hours), "alt s",2])
#     print("added")

# def closeShifts():
#     send(["alt x", 0.5])

def is_float(element: any) -> bool:
    #If you expect None to be passed:
    if element is None:
        return False
    try:
        float(element)
        return True
    except ValueError:
        return False
    
allShifts = getTimes(file)

# openShifts()
# print(allShifts)
# for shift in allShifts:
#     print(shift)
#     addShift(250, shift[0], shift[1], shift[2], datetime.datetime.today().strftime('%m/%d/%Y'))

# closeShifts()


def addshift(employeeID, operation, sequence, hours):
    send(["enter", 0.5])
    send(["#"+str(employeeID),  "enter",0.5, "#"+str(operation),  "enter",0.5, "#"+str(sequence), "enter 8", "#"+str(hours), "enter", "alt s", 1, "enter", 1, "alt x", 1])

def addAllShifts(shifts):
    
    for shift in shifts:
        openTASProgram("WOF")
        addshift(employeeID, shift[0], shift[1], shift[2])

addAllShifts(allShifts)

print("Shifts Added", "Shifts have been added to the system.")