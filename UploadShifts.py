from PlayActions import send_keys as send

from tkinter import filedialog, simpledialog, messagebox
import os
import openpyxl

def getInputs():
    file = filedialog.askopenfilename()
    if file == "":
        os._exit(0)
    
    return file

file = getInputs()

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


def openShifts():
    send(["focus EVO ~ ERP", 0.5])
    send(["alt m m d m", 12])
    send(["tab", "right 2", 0.5, "enter", 0.5])

def addShift(employeeID, operation, sequence, hours, date = ""):
    if date == "":
        send(["enter"])
    else:
        send(["#"+str(date), "enter"])
    
    send(["#"+str(employeeID), "enter", "#"+str(operation), "enter", "#"+str(sequence), "enter 4", "#"+str(hours), "alt s",2])
    print("added")

def closeShifts():
    send(["alt x", 0.5])

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
openShifts()
print(allShifts)
for shift in allShifts:
    print(shift)
    addShift(214, shift[0], shift[1], shift[2], "010119")

closeShifts()

print("Shifts Added", "Shifts have been added to the system.")