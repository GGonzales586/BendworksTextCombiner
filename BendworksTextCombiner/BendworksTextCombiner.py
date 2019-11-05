#! python3
# BendworksTextCombiner.py - A program to combine Greenlee bend reports into one line in an excel workbook for less printing and a cleaner format.

import os, sys, openpyxl, getpass, time
from openpyxl.styles import colors, fonts, alignment, borders
from openpyxl.utils.cell import column_index_from_string
from openpyxl import worksheet
from datetime import datetime, date
from copy import copy

def sortByIndex(List, index):
    List.sort(key = lambda i: i[index])
    return List

user = getpass.getuser()
currentYear = datetime.now().year
currentDate = date.today()

while True:
    try:
        os.chdir('V:\\1. VDC Projects\\GreenleeBendReports\\' + user + '\\ListsOfWantedBends')
        jobNum = str(input('Enter the job number:\n\n'))   # This is used to search the projects folder and get the specific file path name to save the combined file to.
        goodListName = input("\nEnter the name of the txt file with the Mark id's of wanted bend reports to combine. \n(This should be the exported schedule from Revit.)\n\n")
        startTime = time.time()
        goodListName = goodListName + '.txt'
        goodListFile = open(goodListName, 'rb').read()
        break
    except:
        print("\nFile couldn't be opened!\nCheck your spelling.\nMake sure your txt file is in the \"ListsOfWantedBends\" folder\n")

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
        print('\nJob not found in VDC\Projects folder!\nRestartand double check to see if job number exists in the VDC projects folder.\n')
        exit(0)

# This portion creates the excel workbook, reads the wantedFiles values and appends the correct information to the workbook

bendWB = openpyxl.Workbook()
combinedSheet = bendWB.active
combinedSheet.title = 'Combined Bends'
labelSheet = bendWB.create_sheet('Labels')

paramHeaderList = ['Conduit Type', 'Conduit Size', 'Pipe Id', 'Num. of Bends', 'Cut Mark 1', 'Bend Marks', 'Cut Mark 2', 'Bend Rotation', 'Bend Angle', 'Conc. Bends', 'Error Code']
linesToRead = [6, 7, 10, 15, 29, 16, 30, 18, 17, 23, 33]
searchPath = ('V:\\1. VDC Projects\\GreenleeBendReports\\' + user + '\\BendWorksExports')
goodListName = goodListName.replace('.txt', '')
combinedFilePath = targetDir + '\\Combined' + goodListName + '.xlsx'

headerAlignment = alignment.Alignment(horizontal = 'center', vertical = 'center', wrap_text = True)
headerFont = fonts.Font(size = 14, bold = True)

sheetHeaderList = {'Job Name': jobName, 'Job Number': jobNum, 'Area': goodListName, 'Submitted by': user}
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

totalNumOfBends = str(len(masterValues))

listsToRemove = []
concentricBends = []
for sublist in masterValues:  
    if sublist[3] == '0' and sublist[13] == '0':
        listsToRemove.append(masterValues.index(sublist))
    elif sublist[3] == '0' and sublist[13] != '0':
        concentricBends.append(sublist[2])
for index in reversed(listsToRemove):
    masterValues.pop(index)

sortedList = sortByIndex(masterValues, 1)

bendCountSize1 = 0      #bendCount variables for trackin number of bends for the PFS tracking sheet.
bendSize1Param = ['1/2', '3/4', '1', '1 1/4', '1 1/2']
bendCountSize2 = 0
bendSize2Param = ['2', '2 1/2']
bendCountSize3 = 0
bendSize3Param = ['3', '3 1/2']
bendCountSize4 = 0
bendSize4Param = ['4']

for lists in sortedList:
    if lists[1] in bendSize1Param:
        bendCountSize1 += 1
    elif lists[1] in bendSize2Param:
        bendCountSize2 +=1
    elif lists[1] in bendSize3Param:
        bendCountSize3 += 1
    elif lists[1] in bendSize4Param:
        bendCountSize4 += 1
        
for data in sortedList:
    rowCounter +=2
    colCounter = 1
    for elem in data:
        combinedSheet.cell(row = rowCounter, column = colCounter).value = str(elem)
        colCounter +=1

labelColCounter = 1
labelTitles = ['JobName', 'JobNumber', 'Area', 'Type', 'Size', 'Bend Id']
for item in labelTitles:
    labelSheet.cell(column=labelColCounter, row=1, value=item)
    labelColCounter +=1

commonLabels = [jobName, jobNum, goodListName]
labelRowCounter = 1
for lists in sortedList:
    labelRowCounter += 1
    for number in range(len(labelTitles)):
        labelsColCounter = 1
        for items in commonLabels:
            labelSheet.cell(column=labelsColCounter, row=labelRowCounter, value=items)
            labelsColCounter += 1
        for data in lists[0:3]:
            labelSheet.cell(column=labelsColCounter, row=labelRowCounter, value=data)
            labelsColCounter += 1

# This section applies formatting after all the values are processed

for x in range(4):
    combinedSheet.row_dimensions[x].height = 45

colWidths = {'colWidt10' : ['F', 'G', 'H', 'I', 'J', 'O'], 'colWidth12' : ['E', 'N', 'K'], 'colWidth15' : ['A', 'B', 'C', 'D'], 'colWidth35' : ['L', 'M']}

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
# Need: to do some math to output the offset. (if it is an offset or kick 90.)
# Need: to count conduit sticks for ordering
# Need: to get rid of the first cut mark and subtract it from the other marks
# Ned: to figure out orientation of the kwik fit unicouple

bendWB.save(combinedFilePath)
print('\nThe workbook is saved in the following locaton : \n\n' + combinedFilePath)

PFSFilePath = projectsDirPath + '\\' + str(currentYear) + ' PRE-FAB TRACKING.xlsx'
testFilePath = 'V:\\5. VDC - Training\\7. RESEARCH AND DEVELOPMENT\\' + str(currentYear) + ' PRE-FAB TRACKING TEST.xlsx'

PFSTrackingwb = openpyxl.load_workbook(PFSFilePath)    # Replace file path with - PFSFilePath
trackingSheet = PFSTrackingwb.active
goodListName = goodListName.upper() + ' CONDUIT BENDS'
columnValues = {'A': jobNum, 'B': goodListName, 'C': currentDate, 'D': user + '.py', 'E': 1, 'F': 'SHOP', 'T': bendCountSize1, 'U': bendCountSize2, 'V': bendCountSize3, 'W': bendCountSize4}

rowToWriteIn = 5
for line in trackingSheet.iter_rows(min_row = 5, max_row = trackingSheet.max_row, min_col =2, max_col = 2, values_only = True):   
    if goodListName in str(line):
        print('\nBend area already found in PFS Tracking workbook. No values added.\n\nLabels created! Ready for a mail merge!')
        PFSTrackingwb.save(PFSFilePath)     # Replace file path with - PFSFilePath
        endTime = time.time()
        print('Processed ' + totalNumOfBends + ' bends in a whopping ' + str(round(endTime - startTime, 2)) + ' seconds!!!\n')
        exit(0)
    elif line[0] is None:   
        columnCounter = 1
        trackingSheet.insert_rows(rowToWriteIn)
        for column in range(trackingSheet.min_column, trackingSheet.max_column):
            formattedCell = trackingSheet.cell(row = rowToWriteIn - 1, column = columnCounter)
            newCell = trackingSheet.cell(row = rowToWriteIn, column = columnCounter)
            if formattedCell.has_style:
                newCell._style = copy(formattedCell._style)
                columnCounter += 1            
        for col, trackVal in columnValues.items(): 
            trackingSheet.cell(row = rowToWriteIn, column = column_index_from_string(col)).value = trackVal                    
        break
    else:
        rowToWriteIn +=1



PFSTrackingwb.save(PFSFilePath)    # Replace file path with - PFSFilePath
print('\nValues added to the PFS tracking sheet!\n\nLabels created! Ready for a mail merge!')
endTime = time.time()
print('\nProcessed ' + totalNumOfBends + ' bends in a whopping ' + str(round(endTime - startTime, 2)) + ' seconds!!!')
time.sleep(5)

# Inserting a new row needs to carry the sum range with it.





