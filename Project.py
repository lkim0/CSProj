import openpyxl
import os
from openpyxl import Workbook

inputFile = openpyxl.load_workbook('Input Data on Attendance.xlsx')
book = Workbook()
outputSheet = book.active
outputSheet['A1'] = 'Name of Student'
outputSheet['B1'] = 'Grade Level (11 or 12)'
outputSheet['C1'] = 'Date'
outputSheet['D1'] = 'Tardiness'
outputSheet['E1'] = 'Cutting Classes'
outputSheet['F1'] = 'Unexcused Absence'
book.save('Output - Summary of Attendance(1).xlsx')
inputSheet=inputFile['Sheet1']



for i in range(2, len(inputSheet['A'])+1):
    #print(inputSheet["A"+str(i)].internal_value)
    studentName = inputSheet["A"+str(i)].internal_value
    gradeLvl = inputSheet["B"+str(i)].internal_value
    core = inputSheet["C"+str(i)].internal_value
    elective = inputSheet["D"+str(i)].internal_value
    mathLvl = inputSheet["E"+str(i)].internal_value
    date = inputSheet["F"+str(i)].internal_value
    subject = inputSheet["G"+str(i)].internal_value
    remarks = inputSheet["H"+str(i)].internal_value
    maxim=int(len(outputSheet['A']))
    print(studentName)
    ctr = 0
    for value in range(1,maxim+1):
        print("z")
        ctr+=1
        if outputSheet["A"+str(value)].internal_value == studentName:
            print("y")
            if outputSheet["C"+str(value)].internal_value == date:
                if remarks == "T" or remarks == "t":
                    tardiness = outputSheet["D"+str(value)].internal_value
                    tardiness+=1
                    outputSheet["D"+str(value)]=tardiness
                    break
                elif remarks == "A" or remarks == "a":
                    cuttingClasses = outputSheet["E"+str(value)].internal_value
                    cuttingClasses+= 1
                    outputSheet["E"+str(value)]=cuttingClasses
                    break
        elif  ctr == maxim:
            print(value)
            outputSheet['A'+str(maxim+1)]=studentName
            outputSheet['C'+str(maxim+1)]=date
            if remarks == "T" or remarks == "t":
                outputSheet['D'+str(maxim+1)]=1
            elif remarks == "A" or remarks == "a":
                outputSheet['E'+str(maxim+1)]=1
book.save('Output - Summary of Attendance(1).xlsx')
book.close()
inputFile.close()
