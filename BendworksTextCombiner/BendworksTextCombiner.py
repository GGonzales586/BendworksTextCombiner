#! python3
# BendworksTextCombiner.py - A program to combine Greenlee bend reports into one line in an excel workbook for less printing and a cleaner format.

import os, sys, openpyxl, getpass
from openpyxl.styles import colors, fonts, alignment, borders
from openpyxl import worksheet

def sortByIndex(List, index):
    List.sort(key = lambda i: i[index])
    return List

user = getpass.getuser()

while True:
    try:
        os.chdir('V:\\1. VDC Projects\\GreenleeBendReports\\' + user + '\\ListsOfWantedBends')
        jobNum = str(input('Enter the job number: '))   # This is used to search the projects folder and get the specific file path name to save the combined file to.
        goodListName = input("Enter the name of the txt file with the Mark id's of wanted bend reports to combine. \n(This should be the exported schedule from Revit.)\n")
        goodListName = goodListName + '.txt'
        goodListFile = open(goodListName, 'rb').read()
        break
    except:
        print("File couldn't be opened!\nCheck your spelling.\nMake sure your txt file is in the \"ListsOfWantedBends\" folder")

fileText = goodListFile.decode('utf16')     #Revit schedules need to be decoded to be able to read them.
wantedFiles = fileText.splitlines()
listOfFileNames = []

for val in wantedFiles:
    val = val.replace('"', '')
    val = val + '.txt'
    listOfFileNames.append(val)      # This list will used further down to open each file with the names in this list and read the wnated lines for combining.

# This portion sets up the file paths to be saved to the correct project folder within the PFS folder.

projectsDirPath = ('V:\\1. VDC Projects')
projectDirNames = os.listdir('V:\\1. VDC Projects')
targetDir = ('\\3- PFS\\GreenleeUsedBendReports')

jobFound = False
while not jobFound:
    if jobFound == True:
        break
    try:
        for directory in projectDirNames:
            if jobFound == True:
                break
            else:
                if jobNum in directory:
                    jobDirPath = (projectsDirPath + '\\' + directory)
                    jobName = str(jobDirPath.split(' - ')[1])
                    jobFound = True
        targetDir = jobDirPath + targetDir
    except:
        print('\nJob not found in VDC\Projects folder!\nRestartand double check to see if job number exists in the VDC projects folder.')
        input('Hit "Enter" key to close.')
        sys.exit()

# This portion creates the excel workbook, reads the wantedFiles values and appends the correct information to the workbook

bendWB = openpyxl.Workbook()
combinedSheet = bendWB.active
combinedSheet.title = 'Combined Bends'

paramHeaderList = ['Conduit Type', 'Conduit Size', 'Pipe Id', 'Num. of Bends', 'Cut Mark 1', 'Bend Marks', 'Cut Mark 2', 'Bend Rotation', 'Bend Angle', 'Conc. Bends', 'Error Code']
linesToRead = [6, 7, 10, 15, 29, 16, 30, 18, 17, 23, 33]
searchPath = ('V:\\1. VDC Projects\\GreenleeBendReports\\' + user + '\\BendWorksExports')
goodListName = goodListName.replace('.txt', '')
combinedFilePath = targetDir + '\\Combined' + goodListName + '.xlsx'

headerAlignment = alignment.Alignment(horizontal = 'center', vertical = 'center', wrap_text = True)
headerFont = fonts.Font(size = 14, bold = True)

sheetHeaderList = {'Job Name': jobName, 'Job Number': jobNum, 'Submitted by': user}
colCounter = 1
for key, vals in sheetHeaderList.items():
    combinedSheet.cell(row = 1, column = colCounter).value = key
    combinedSheet.cell(row = 2, column = colCounter).value = vals
    combinedSheet.cell(row = 1, column = colCounter).alignment = headerAlignment
    combinedSheet.cell(row = 1, column = colCounter).font = headerFont
    combinedSheet.cell(row = 2, column = colCounter).alignment = headerAlignment
    combinedSheet.cell(row = 2, column = colCounter).font = headerFont
    colCounter += 1

combinedSheet.merge_cells('F3:J3')

colCounter = 1
for vals in paramHeaderList:
    if vals == 'Bend Marks':
        combinedSheet.cell(row = 3, column = 6).value = vals
        combinedSheet.cell(row = 3, column = 6).alignment = headerAlignment
        combinedSheet.cell(row = 3, column = 6).font = headerFont
        colCounter += 5
    else:
        combinedSheet.cell(row = 3, column = colCounter).value = vals
        combinedSheet.cell(row = 3, column = colCounter).alignment = headerAlignment
        combinedSheet.cell(row = 3, column = colCounter).font = headerFont
        colCounter +=1

masterValues = []
rowCounter = 3
for file in os.listdir(searchPath):
    if file in listOfFileNames:        
        openFile = open(searchPath + '\\' + file).read()
        splitFile = openFile.splitlines()
        wantedValues = []
        if jobNum in str(splitFile[0]):
            for num in linesToRead:
                value = splitFile[num]
                if num == 10:
                    splitValue = value.split('       ')
                    val2 = splitValue[1].strip()
                    wantedValues.append(val2)
                elif num == 16:
                    splitValue = value.split(':')
                    del splitValue[0]
                    stringValue = ''
                    for v in splitValue:
                        stringValue += v
                    splitValue = stringValue.split('"')
                    splitValue.pop()
                    for v in splitValue:
                        val = v.strip()
                        wantedValues.append(val)
                else:
                    splitValue = value.split(':')
                    val2 = splitValue[1].strip()
                    wantedValues.append(val2)
        masterValues.append(wantedValues)

sortedList = sortByIndex(masterValues, 1)
        
for data in sortedList:
    rowCounter +=2
    colCounter = 1
    for elem in data:
        combinedSheet.cell(row = rowCounter, column = colCounter).value = str(elem)
        colCounter +=1

# This section applies formatting after all the values are processed

for x in range(4):
    combinedSheet.row_dimensions[x].height = 45

colWidths = {'colWidt10' : ['F', 'G', 'H', 'I', 'J', 'O'], 'colWidth12' : ['D', 'E', 'N', 'K'], 'colWidth15' : ['A', 'B', 'C'], 'colWidth35' : ['L', 'M']}

for key, value in colWidths.items():
    if '10' in key:
        for item in value:
            combinedSheet.column_dimensions[item].width = 10
    if '12' in key:
        for item in value:
            combinedSheet.column_dimensions[item].width = 12
    if '15' in key:
        for item in value:
            combinedSheet.column_dimensions[item].width = 15
    if '35' in key:
        for item in value:
            combinedSheet.column_dimensions[item].width = 35

valueRange = combinedSheet.max_row + 1
pipeIdColor = fonts.Font(color = 'FF9900')
cutMarkColor = fonts.Font(color = 'FF0000')
bendMarkColor = fonts.Font(color = '0000FF')
bendRotColor = fonts.Font(color = 'FF1493')
bendAngleColor = fonts.Font(color = '008000')

for cell in range(5 , valueRange):
    combinedSheet.cell(row = cell, column = 3).font = pipeIdColor
    combinedSheet.cell(row = cell, column = 5).font = cutMarkColor    
    combinedSheet.cell(row = cell, column = 6).font = bendMarkColor
    combinedSheet.cell(row = cell, column = 7).font = bendMarkColor
    combinedSheet.cell(row = cell, column = 8).font = bendMarkColor
    combinedSheet.cell(row = cell, column = 9).font = bendMarkColor
    combinedSheet.cell(row = cell, column = 10).font = bendMarkColor
    combinedSheet.cell(row = cell, column = 11).font = cutMarkColor
    combinedSheet.cell(row = cell, column = 12).font = bendRotColor
    combinedSheet.cell(row = cell, column = 13).font = bendAngleColor

# Need: to format printing 10 11x17 landscape
# Need: to format border
# Need: to print header at the top of each 11x17 sheet


bendWB.save(combinedFilePath)
print('The workbook is saved in the following locaton : \n\n' + combinedFilePath)


