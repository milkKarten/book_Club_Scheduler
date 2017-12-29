'''Created by Seth Karten, 2017'''
'''
Creates REAL groups based on timeslots and REAL leader info
IMPORTANT: Include only firstName, lastName, REAL leader,
		   and availability fields in excel spreadsheet
'''
import openpyxl
import random
import math
import time

'''Set to name and directory of excel file'''
wbName = 'rutgers_engineers_assessing_literature_real_program_availability_form.xlsx'
wb = openpyxl.load_workbook(wbName)
sheet = wb.get_sheet_by_name('Sheet1')

numEntries = sheet.max_row
maxColumns = 45
firstNameColumn = 'A'
lastNameColumn = 'B'
isLeaderColumn = 'C'
blankCell = sheet['A1'].value
numLeaders = 0
realLeaders = []
groupList = []
realLeaderIndicie = []

'''Label top columns by number'''
for i in range(4, sheet.max_column + 1):
	sheet.cell(row = 1, column = i).value = i

'''Get rid of duplicates and mark down REAL leaders'''
for person in range(3, numEntries + 1):
	duplicate = False
	firstName = sheet[firstNameColumn + str(person)].value
	lastName = sheet[lastNameColumn + str(person)].value
	for possibleDuplicate in range(person + 1, numEntries + 1):
		curFirstName = sheet[firstNameColumn + str(possibleDuplicate)].value
		curLastName = sheet[lastNameColumn + str(possibleDuplicate)].value
		if firstName == curFirstName and lastName == curLastName:
			sheet[firstNameColumn + str(person)] = blankCell
			sheet[lastNameColumn + str(person)] = blankCell
			duplicate = True
	if not duplicate:
		# Manage number of leaders and create groups for
		# leaders with leaders as only member so far
		if sheet[isLeaderColumn + str(person)].value == 'Yes':
			numLeaders += 1
			groupName = 'Group: ' + str(firstName) + ' ' + str(lastName)
			realLeaders.append(str(firstName) + ' ' + str(lastName))
			realLeaderIndicie.append(person)
			groupList.append([[groupName],[person],1, 1]) # names, row nums of names, meeting time column, mode
			#exec('group%d = [[member],[person],[0]]'%numLeaders) # name, row num, meeting time column
#print('Number of Leaders: ' + str(numLeaders))
#print(groupList)

'''Sort Availabilities'''
modes = [] # Ordered list of modes from largest value to smallest
for columnNum in range(4, maxColumns + 1):
	numAvailable = 0
	for rowNum in range(4, numEntries + 1):
		isAvailable = sheet.cell(row = rowNum, column = columnNum)
		if isAvailable.value == 'Available':
			numAvailable += 1
	for i in range(0, len(modes) + 1, 1):
		if i == len(modes):
			modes.append([numAvailable, columnNum])
		elif modes[i][0] < numAvailable:
			tempMode = modes[i][0]
			tempColumnNum = modes[i][1]
			modes[i] = [numAvailable, columnNum]
			numAvailable = tempMode
			columnNum = tempColumnNum
	if len(modes) == 0:
		modes.append([numAvailable, columnNum])
#print('Modes: ' + str(modes))
#print(len(modes))

'''Assign Leaders to mode meeting times'''
finishedLeaders = []
for modeIndicie in range(numLeaders - 1, -1, -1):
	vertice = 0
	for currentLeader in realLeaders:
		if currentLeader in finishedLeaders:
			#print(str(currentLeader) + ' is done')
			vertice += 1
			continue
		columnIndicie = modes[modeIndicie][1]
		isAvailable = sheet.cell(row = groupList[vertice][1][0], column = columnIndicie)
		if isAvailable.value == 'Available':
			groupList[vertice][3] = modes[modeIndicie][0]
			groupList[vertice][2] = columnIndicie
			groupList[vertice][0] = [groupList[vertice][0][0] + ' ' + str(sheet.cell(row = 3, column = columnIndicie).value)]
			#print(groupList[vertice][0])
			#print(vertice)
			#print(currentLeader)
			finishedLeaders.append(currentLeader)
			break
		vertice += 1
# In case leaders need less popular times
#print(realLeaders)
#print(finishedLeaders)
if len(finishedLeaders) != numLeaders:
	#print('Less popular times needed...')
	for modeIndicie in range(numLeaders, len(modes)):
		vertice = 0
		if len(finishedLeaders) == numLeaders:
			break
		for currentLeader in realLeaders:
			if currentLeader in finishedLeaders:
				vertice += 1
				continue
			columnIndicie = modes[modeIndicie][1]
			isAvailable = sheet.cell(row = groupList[vertice][1][0], column = columnIndicie)
			#exec('isAvailable = sheet.cell(row = group%d[1][vertice], column = modes[modeIndicie][1])'%currentLeader)
			if isAvailable.value == 'Available':
				groupList[vertice][3] = modes[modeIndicie][0]
				groupList[vertice][2] = columnIndicie
				groupList[vertice][0] = [groupList[vertice][0][0] + ' ' + str(sheet.cell(row = 3, column = columnIndicie).value)]
				#exec('group%d[2][vertice] = modes[modeIndicie][1]'%currentLeader)
				finishedLeaders.append(currentLeader)
				#(groupList[vertice][0])
				break
			vertice += 1
#Leaders cannot make any times... time to start switching
#print(groupList)
if len(finishedLeaders) != numLeaders or len(groupList) != numLeaders:
	print('not all leaders have groups')
	for leader in range(0, len(realLeaders)):
		if realLeaders[leader] not in finishedLeaders:
			firstName = sheet.cell(row = realLeaderIndicie[leader], column = 1).value
			lastName = sheet.cell(row = realLeaderIndicie[leader], column = 2).value
			print('(' + str(leader) + '): ' + str(firstName) + ' ' + str(lastName) + ' did not create group')
#print(finishedLeaders)
'''Order leaders by mode'''
def orderByMode(groupList):
	orderedGroupList = [groupList[0]]
	for k in range(1, len(groupList)):
		compare = groupList[k][3]
		compareTerm = groupList[k]
		for j in range(0, len(orderedGroupList)):
			if compare > orderedGroupList[j][3]:
				#print(str(compare) + '\t' + '>' + '\t' + str(orderedGroupList[j][3]))
				compareTermTemp = orderedGroupList[j]
				compare = orderedGroupList[j][3]
				orderedGroupList[j] = compareTerm
				compareTerm = compareTermTemp
			if j == len(orderedGroupList) - 1:
				#print(str(compare) + '\t' + '<' + '\t' + str(orderedGroupList[j][3]))
				orderedGroupList.append(compareTerm)
				break
	return orderedGroupList

groupList = orderByMode(groupList)
#print(groupList)
'''Assign people to group with least common mode that works'''
finishedMembers = finishedLeaders
for member in range(4, numEntries + 1):
	firstName = sheet.cell(row = member, column = 1).value
	lastName = sheet.cell(row = member, column = 2).value
	name = str(firstName) + ' ' + str(lastName)
	if firstName == blankCell or name in finishedMembers:
		continue
	for currentLeader in range(len(groupList)-1, -1, -1):
		if firstName == blankCell or member in finishedMembers:
			continue
		isAvailable = sheet.cell(row = member, column = groupList[currentLeader][2])
		if isAvailable.value == 'Available':
			#print(groupList)
			groupList[currentLeader][0].append(name)
			groupList[currentLeader][1].append(member)
			#exec('group%d[0].append(firstName + ' ' + lastName)'%currentLeader)
			finishedMembers.append(name)
			break
	if name not in finishedMembers:
		print('(' + str(member) + '): ' + str(firstName) + ' ' + str(lastName) + ' did not join group')

'''Even number of group members'''
totalNumMembers = len(finishedMembers)
#print(totalNumMembers)
#print(len(groupList))
membersPerGroup = math.floor(totalNumMembers / len(groupList))
#print(membersPerGroup)
startingTime = time.time()
for group in range(0, len(groupList) - 1):
	while(len(groupList[group][0]) < membersPerGroup + 1):
		secs = time.time()
		if (secs - startingTime) > 60.0:
			break
		transferFrom = random.randint(-1, len(groupList) - 1)
		if len(groupList[transferFrom][0]) > membersPerGroup:
			availSwitchers = []
			for member in range(1, len(groupList[transferFrom][1])):
				#print(groupList[transferFrom][1][member])
				if sheet.cell(row = groupList[transferFrom][1][member], column = groupList[transferFrom][2]).value == 'Available':
					availSwitchers.append(groupList[transferFrom][0][member])
			if len(availSwitchers) > 0:
				transferMember = random.choice(availSwitchers)
				transferIndex = groupList[transferFrom][0].index(transferMember)
				#print(transferIndex)
				transferMemberRow = groupList[transferFrom][1][transferIndex]
				groupList[transferFrom][0].remove(transferMember)
				groupList[transferFrom][1].remove(transferMemberRow)
				groupList[group][0].append(transferMember)


#for i in range(0, len(groupList)):
	#print(groupList[i])
#Work in-progress here

'''Save file with updated info'''
if sheet['C1'].value == blankCell:
	wb.create_sheet('Groups')
	sheet['C1'].value = 'Clear this cell if Groups sheet is deleted'
else:
	sheetName = wb.get_sheet_by_name('Groups')
	wb.remove_sheet(sheetName)
	wb.create_sheet('Groups')
groupSheet = wb.get_sheet_by_name('Groups')
for i in range(1, len(groupList) + 1):
	cell = str(sheet.cell(row = 3, column = i + 3))
	cell = cell[15:]
	if cell[2] == '>':
		cell = cell[0:2]
	groupSheet.cell(row = 1, column = i).value = sheet[cell].value
	numMembers = len(groupList[i - 1][0])
	#exec('numMembers = len(group%d[0])'%i)
	for j in range(0, numMembers):
		groupSheet.cell(row = j + 1, column = i).value = groupList[i - 1][0][j]
		#exec('groupSheet.cell(row = j, column = i).value = group%d[0][j]'%i)
wb.save(wbName)
