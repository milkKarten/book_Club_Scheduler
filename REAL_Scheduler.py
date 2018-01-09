#! /usr/bin/env python
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

class schedule:
	def __init__(self):
		self.groupList = [] # names, row nums of names, meeting time column, mode
		self.unassignedMembers = [] # names, row nums of names
		self.membersNotAssigned = False
		self.finishedMembers = []

'''Assign people to group with most common mode that works'''
def assignGroups(this):
	for member in range(4, numEntries + 1): # Corresponds to row on excel sheet
		firstName = sheet.cell(row = member, column = 1).value
		lastName = sheet.cell(row = member, column = 2).value
		name = str(firstName) + ' ' + str(lastName)
		if firstName == blankCell or name in this.finishedMembers:
			continue
		for currentLeader in range(len(this.groupList) - 1, -1, -1):
			if firstName == blankCell or member in this.finishedMembers:
				continue
			isAvailable = sheet.cell(row = member, column = this.groupList[currentLeader][2])
			if isAvailable.value == 'Available':
				#print(this.groupList)
				this.groupList[currentLeader][0].append(name)
				this.groupList[currentLeader][1].append(member)
				#exec('group%d[0].append(firstName + ' ' + lastName)'%currentLeader)
				this.finishedMembers.append(name)
				break
		if name not in this.finishedMembers:
			this.membersNotAssigned = True
			if len(this.unassignedMembers) == 0:
				this.unassignedMembers = [[name], [member]]
			else:
				this.unassignedMembers[0].append(name)
				this.unassignedMembers[1].append(member)
			print('(' + str(member) + '): ' + str(firstName) + ' ' + str(lastName) + ' did not join group')


'''Order leaders by mode using insertion sort'''
def orderByMode(this):
	orderedGroupList = [this.groupList[0]]
	for k in range(1, len(this.groupList)):
		compare = this.groupList[k][3]
		compareTerm = this.groupList[k]
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

this = schedule()

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
			groupName = 'Group: ' + str(firstName) + ' ' + str(lastName) + ' -'
			realLeaders.append(str(firstName) + ' ' + str(lastName))
			realLeaderIndicie.append(person)
			this.groupList.append([[groupName],[person],1, 1]) # names, row nums of names, meeting time column, mode
			#exec('group%d = [[member],[person],[0]]'%numLeaders) # name, row num, meeting time column
#print('Number of Leaders: ' + str(numLeaders))
#print(this.groupList)

'''Sort Availabilities'''
modes = [] # Ordered list of modes from largest value to smallest; [numAvailable, columnNum]
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
		isAvailable = sheet.cell(row = this.groupList[vertice][1][0], column = columnIndicie)
		if isAvailable.value == 'Available':
			this.groupList[vertice][3] = modes[modeIndicie][0]
			this.groupList[vertice][2] = columnIndicie
			this.groupList[vertice][0] = [this.groupList[vertice][0][0] + ' ' + str(sheet.cell(row = 3, column = columnIndicie).value)]
			#print(this.groupList[vertice][0])
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
			isAvailable = sheet.cell(row = this.groupList[vertice][1][0], column = columnIndicie)
			#exec('isAvailable = sheet.cell(row = group%d[1][vertice], column = modes[modeIndicie][1])'%currentLeader)
			if isAvailable.value == 'Available':
				this.groupList[vertice][3] = modes[modeIndicie][0]
				this.groupList[vertice][2] = columnIndicie
				this.groupList[vertice][0] = [this.groupList[vertice][0][0] + ' ' + str(sheet.cell(row = 3, column = columnIndicie).value)]
				#exec('group%d[2][vertice] = modes[modeIndicie][1]'%currentLeader)
				finishedLeaders.append(currentLeader)
				#(this.groupList[vertice][0])
				break
			vertice += 1
#Leaders cannot make any times... time to start switching
#print(this.groupList)
if len(finishedLeaders) != numLeaders or len(this.groupList) != numLeaders:
	print('not all leaders have groups')
	for leader in range(0, len(realLeaders)):
		if realLeaders[leader] not in finishedLeaders:
			firstName = sheet.cell(row = realLeaderIndicie[leader], column = 1).value
			lastName = sheet.cell(row = realLeaderIndicie[leader], column = 2).value
			print('(' + str(leader) + '): ' + str(firstName) + ' ' + str(lastName) + ' did not create group')
#print(finishedLeaders)

this.groupList = orderByMode(this)
#print(this.groupList)
'''Assign people to group with most common mode that works'''
this.finishedMembers = finishedLeaders
assignGroups(this)

'''Switch leaders to new timeslot if a member does not fit in a group'''
orderedModes = [[modes[0][0]], [modes[0][1]]]
while len(this.unassignedMembers) > 0:
	newMembersToAssign = []
	for i in range(1, len(modes)):
		orderedModes[0].append(modes[i][0]) # Mode
		orderedModes[1].append(modes[i][1]) # timeslot column
	if this.membersNotAssigned:
		for k in range(0, len(this.unassignedMembers[0])):
			movedMember = False
			for mode in range(0, len(orderedModes[0])):
				inList = False
				for j in range(0, len(this.groupList)):
					if orderedModes[0][mode] == this.groupList[j][3]:
						inList = True
						break
				if inList:
					continue
				member = this.unassignedMembers[1][k]
				isAvailable = sheet.cell(row = member, column = orderedModes[1][mode])
				if isAvailable.value == 'Available':
					for group in range(0, len(this.groupList)):
						allMembersCanMove = True
						for rowNum in this.groupList[group][1]:
							isMemberAvailable = sheet.cell(row = rowNum, column = orderedModes[1][mode])
							if isMemberAvailable != 'Available':
								allMembersCanMove = False
								break
						if allMembersCanMove:
							print('WOWWWWWW')
							print('Moving Group')
							this.groupList[group][2] = orderedModes[1][mode]
							this.groupList[group][3] = orderedModes[0][mode]
							this.groupList[group][0].append(this.unassignedMembers[0][k])
							this.groupList[group][1].append(this.unassignedMembers[1][k])
							this.unassignedMembers.remove(this.unassignedMembers[0][k])
							this.unassignedMembers.remove(this.unassignedMembers[1][k])
					if not allMembersCanMove:
						cannotMove = []
						canMove = False
						while True:
							group = random.randint(0, len(this.groupList) - 1)
							if group not in cannotMove:
								isAvailable = sheet.cell(row = this.groupList[group][1][0], column = orderedModes[1][mode])
								if isAvailable.value == 'Available':
									canMove = True
									break
								else:
									canMove = False
									cannotMove.append(group)
							if len(cannotMove) == len(this.groupList):
								print('Cannot switch group: ' + str(sheet.cell(row = member, column = 1).value) + ' ' + str(sheet.cell(row = member, column = 2).value) + ' must pick from an available timeslot manually')
								break
						if canMove:
							print('Moving')
							movedMember = True
							if len(newMembersToAssign) == 0:
								newMembersToAssign = [this.groupList[group][0][1:], this.groupList[group][1][1:]]
							else:
								for mem in range(1, len(this.groupList[group][0])):
									newMembersToAssign[0].append(this.groupList[group][0][mem])
									newMembersToAssign[1].append(this.groupList[group][1][mem])
							this.groupList[group][0] = [this.groupList[group][0][0], this.unassignedMembers[0][k]]
							this.groupList[group][1] = [this.groupList[group][1][0], this.unassignedMembers[1][k]]
							this.groupList[group][2] = orderedModes[1][mode]
							this.groupList[group][3] = orderedModes[0][mode]
							substringIndex = this.groupList[group][0][0].index('-') + 1
							subString = this.groupList[group][0][0][:substringIndex]
							this.groupList[group][0][0] = subString + ' ' + str(sheet.cell(row = 3, column = this.groupList[group][2]).value)

				if movedMember:
					break
		#Assign newMembersToAssign to groups
		newFinishedMembers = []
		this.unassignedMembers = []
		for member in range(0, len(newMembersToAssign[1])):
			name = newMembersToAssign[0][member]
			if name in newFinishedMembers:
				continue
			for currentLeader in range(0, len(this.groupList)):
				rowNum = newMembersToAssign[1][member]
				columnNum = this.groupList[currentLeader][2]
				isAvailable = sheet.cell(row = rowNum, column = columnNum)
				if isAvailable.value == 'Available':
					this.groupList[currentLeader][0].append(name)
					this.groupList[currentLeader][1].append(newMembersToAssign[1][member])
					newFinishedMembers.append(name)
					break
			if name not in newFinishedMembers:
				this.membersNotAssigned = True
				if len(this.unassignedMembers) == 0:
					this.unassignedMembers = [[name], [newMembersToAssign[1][member]]]
				else:
					this.unassignedMembers[0].append(name)
					this.unassignedMembers[1].append(newMembersToAssign[1][member])
				print(str(name) + ' did not join group')
				#print(this.unassignedMembers)

'''Assign each group average number of group members'''
totalNumMembers = len(this.finishedMembers)
#print(totalNumMembers)
#print(len(this.groupList))
membersPerGroup = math.floor(totalNumMembers / len(this.groupList))
#print(membersPerGroup)
startingTime = time.time()
while time.time() - startingTime < 5.0:
	for group in range(0, len(this.groupList)):
		while (len(this.groupList[group][0]) < membersPerGroup + 1):
			secs = time.time()
			if (secs - startingTime) > 1.0:
				break
			transferFrom = random.randint(-1, len(this.groupList) - 1)
			if len(this.groupList[transferFrom][0]) > membersPerGroup:
				availSwitchers = [] # names
				availSwitchersIndicie = [] # row index
				for member in range(1, len(this.groupList[transferFrom][1])):
					#print(this.groupList[transferFrom][1][member])
					if sheet.cell(row = this.groupList[transferFrom][1][member], column = this.groupList[group][2]).value == 'Available':
						availSwitchers.append(this.groupList[transferFrom][0][member])
						availSwitchersIndicie.append(this.groupList[transferFrom][1][member])

				if len(availSwitchers) > 0:
					transferMember = random.choice(availSwitchers)
					transferIndex = availSwitchers.index(transferMember)
					#print(transferIndex)
					transferMemberRow = availSwitchersIndicie[transferIndex]
					this.groupList[transferFrom][0].remove(transferMember)
					this.groupList[transferFrom][1].remove(transferMemberRow)
					this.groupList[group][0].append(transferMember)
					this.groupList[group][1].append(transferMemberRow)
		while (len(this.groupList[group][0]) > membersPerGroup + 2):
			secs = time.time()
			if (secs - startingTime) > 3.0:
				break
			transferTo = random.randint(0, len(this.groupList) - 1)
			if transferTo == group:
				continue
			#print('transfer to ' + str(transferTo))
			availSwitchers = [] # names
			availSwitchersIndicie = [] # row index
			for member in range(1, len(this.groupList[group][1])):
				#print(this.groupList[transferTo][1][member])
				#print(str(this.groupList[group][1][member]) + ' ' + str(this.groupList[transferTo][2]))
				if sheet.cell(row = this.groupList[group][1][member], column = this.groupList[transferTo][2]).value == 'Available':
					availSwitchers.append(this.groupList[group][0][member])
					availSwitchersIndicie.append(this.groupList[group][1][member])
				if (len(this.groupList[group][0]) - len(availSwitchers)) <= (membersPerGroup + 2):
					break
			if len(availSwitchers) > 0:
				transferMember = random.choice(availSwitchers)
				transferIndex = availSwitchers.index(transferMember)
				#print(transferIndex)
				transferMemberRow = availSwitchersIndicie[transferIndex]
				this.groupList[transferTo][0].append(availSwitchers[transferIndex])
				this.groupList[transferTo][1].append(transferMemberRow)
				this.groupList[group][0].remove(availSwitchers[transferIndex])
				this.groupList[group][1].remove(transferMemberRow)

#for i in range(0, len(this.groupList)):
	#print(this.groupList[i])

'''Check Groups'''
for group in this.groupList:
	for member in range(0, len(group[1])):
		isAvailable = sheet.cell(row = group[1][member], column = group[2])
		if isAvailable.value != 'Available':
			print(str(group[0][member]) + ' did not join a proper timeslot at ' + str(sheet.cell(row = 3, column = group[2]).value))

'''Save file with updated info'''
if sheet['C1'].value == blankCell:
	wb.create_sheet('Groups')
	sheet['C1'].value = 'Clear this cell if Groups sheet is deleted'
else:
	sheetName = wb.get_sheet_by_name('Groups')
	wb.remove_sheet(sheetName)
	wb.create_sheet('Groups')
groupSheet = wb.get_sheet_by_name('Groups')
for i in range(1, len(this.groupList) + 1):
	cell = str(sheet.cell(row = 3, column = i + 3))
	cell = cell[15:]
	if cell[2] == '>':
		cell = cell[0:2]
	groupSheet.cell(row = 1, column = i).value = sheet[cell].value
	numMembers = len(this.groupList[i - 1][0])
	#exec('numMembers = len(group%d[0])'%i)
	for j in range(0, numMembers):
		groupSheet.cell(row = j + 1, column = i).value = this.groupList[i - 1][0][j]
		#exec('groupSheet.cell(row = j, column = i).value = group%d[0][j]'%i)
wb.save(wbName)
