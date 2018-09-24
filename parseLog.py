#Get result from log file
#Nghia Tran
#September - 09 - 2018

import re
import xlsxwriter
import glob
import os

#Create result
def createTextResult(fileName):    
    sourceFile = open(fileName, "r")   #source file
    resultFile = open("Result.txt", "w")

    for line in sourceFile:
        if ("UCAPPSSE" in line or "UCAPM" in line or "Info" in line) and (not "pkg" in line) and (not "Alpha" in line) and (not "Beta" in line):
            if ". " in line:
                newLine1 = line.split(". ")[1].replace("(", "<").replace(")",">")
                newLine2 =  re.sub('<.*?>', '', newLine1)
                resultFile.write(newLine2)
            else:
                resultFile.write(line)
    resultFile.write("Done..............")
    sourceFile.close()
    resultFile.close()

#Parse result
def getNumberOfTestCase():
    count = 0
    sourceFile = open("Result.txt", "r")
    for line in sourceFile:
        if ("UCAPPSSE" in line or "UCAPM" in line) and not "Info" in line:
            count = count + 1
    sourceFile.close()
    return count

#Index
def getIndexOfTestCase():
    indexArray = []
    for i in range(0,getNumberOfTestCase()):
        indexArray.append(i+1)
    return indexArray

#TestCase name
def getTestCaseArray():
    testCaseArray = []
    sourceFile = open("Result.txt", "r")
    for line in sourceFile:
        if ("UCAPPSSE" in line or "UCAPM" in line) and not "Info" in line:
            tescaseName = line.split(' ')[0]
            testCaseArray.append(tescaseName)
    sourceFile.close()
    return testCaseArray

#Status
def getTestCaseStatus():
    statusArray = []
    sourceFile = open("Result.txt", "r")
    for line in sourceFile:
        if ("UCAPPSSE" in line or "UCAPM" in line) and not "Info" in line:
            if "PASS" in line:
                statusArray.append("Passed")
            else:
                statusArray.append("Failed")
    sourceFile.close()
    return statusArray
    
#Issue
def printLineWithNumber(number):
    sourceFile = open("Result.txt", "r")
    info = ''
    line = sourceFile.readlines()
    for i in range(0,10):
        if "Info" in line[number + i]:
            info += line[number + i]
        else:
            break
    sourceFile.close()
    info = info + "x"
    info = info.replace("\nx","")
    #remove "Info : java.lang.Exception: "
    info = info.replace("Info : java.lang.Exception: ","")
    return info

def getIndexFailedTestCase():
    lookup = 'FAIL      '
    failedArray = []
    with open("Result.txt") as myFile:
        for num, line in enumerate(myFile, 1):
            if lookup in line:
                failedArray.append(num)
    return failedArray

def getIssue():
    issueArray = []
    k = 0
    for i in range(0, getNumberOfTestCase()):
        for j in range(0, len(getIndexFailedTestCase())):
            if getTestCaseStatus()[i] == "Passed":
                issueArray.append("None")
            else:
                issueArray.append(printLineWithNumber(getIndexFailedTestCase()[k]))
                k = k + 1
            break
    return issueArray

#Time
def getTime():
    timeArray = []
    sourceFile = open("Result.txt", "r")
    for line in sourceFile:
        if ("UCAPPSSE" in line or "UCAPM" in line) and not "Info" in line:
            if "PASS" in line:
                time = line.split("PASS      ")[1].replace(" \n", "")
                timeArray.append(time)
            else:
                time = line.split("FAIL      ")[1].replace(" \n", "")
                timeArray.append(time)
    sourceFile.close()
    return timeArray

def createExcelDataFrame():
    print("Creating Excel DataFrame...")
    w, h = 5, getNumberOfTestCase()
    matrix = [[0 for x in range(w)] for y in range(h)]
    indexArray = getIndexOfTestCase()
    testCaseArray = getTestCaseArray()
    testCaseStatusArray = getTestCaseStatus()
    issueArray = getIssue()
    timeArray = getTime()
    
    for i in range(0, getNumberOfTestCase()):
        matrix[i][0] = indexArray[i]
        matrix[i][1] = testCaseArray[i]
        matrix[i][2] = testCaseStatusArray[i]
        matrix[i][3] = issueArray[i]
        matrix[i][4] = timeArray[i]
    print(str(getNumberOfTestCase()) + " rows done...")
    return matrix



#------------------------------------------------------EXCEL-------------------------------------
def CreateExcelFile():
    
    print("\nCreating Excel file...")
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some data headers.
    worksheet.write('A1', 'Number', bold)
    worksheet.write('B1', 'TestCase', bold)
    worksheet.write('C1', 'Status', bold)
    worksheet.write('D1', 'Issue', bold)
    worksheet.write('E1', 'Time', bold)

    # Some data we want to write to the worksheet.
    example = (
        [1,'UCAPPSSE_68_TC52', 'Passed', 'None', 'Today'],
        [2,'UCAPPSSE_68_TC53', 'Passed', 'None', 'Today'],
        [3,'UCAPPSSE_68_TC54', 'Failed', 'Info : java.lang.Exception: ACW30_UnHold;1 on 10.128.224.49 is Fail', 'Today'],
        [4,'UCAPPSSE_68_TC55', 'Passed', 'None', 'Today'],
        [5,'UCAPPSSE_68_TC56', 'Passed', 'None', 'Today'],
    )

    # Start from the first cell below the headers.
    row = 1
    col = 0

    # Iterate over the data and write it out row by row.
    for number, testCase, status, issue, time in (createExcelDataFrame()):
        worksheet.write(row, col,     number)
        worksheet.write(row, col + 1, testCase)
        worksheet.write(row, col + 2, status)
        worksheet.write(row, col + 3, issue)
        worksheet.write(row, col + 4, time)
        row += 1
    print("Created Excel file result.xlsx") 
    workbook.close()

if __name__ =="__main__":
    print("\nTxt file(s): ")
    print (glob.glob('./*.txt'))
    print()
    file = str(input("Input file name (*.txt) : "))
    createTextResult(file)
    print("\nTotal TCs: " + str(getNumberOfTestCase()))
    print("Failed TCs: " + str(len(getIndexFailedTestCase())))
    print("Passed TCs: " + str(getNumberOfTestCase() - len(getIndexFailedTestCase())) + "\n")
    os.system("pause")
    CreateExcelFile()