import xlrd
import xlwt
try:
    ourbook = raw_input("Enter the filename of the Scouting Spreadsheet: ")
except IOError:
    while True:
        try:
            ourbook = raw_input("Enter the filename of the Scouting Spreadsheet: ")
        except IOError:
            ourbook = raw_input("Enter the filename of the Scouting Spreadsheet: ")
workbook = xlrd.open_workbook(ourbook+'.xlsx')
#oursheet = raw_input("Enter the sheet name within the spreadsheet: ")
matchsheet = workbook.sheet_by_name('MATCH SCOUTING')
pitsheet = workbook.sheet_by_name('PIT SCOUTING')
#ourRows = input("Enter the number of filled rows in the spreadsheet: ")
numberOfRowsMatch = input("Enter the number of filled rows in Match sheet: ")
numberOfRowsPit = input("Enter the number of filled rows in Pit sheet: ")
allMatchRowTeamNumbers = []
allMatchRowData = {}
allTeams = []
matchColumns = ['Match Number','Alliance Color','Robot Starting Position','Autonomous Performance','Autonomous Notes','Switch Cube Delivery','Scale Cube Delivery','Exchange Cube Delivery','Maneuverability','Driver Skill','Teleoperation Notes','Endgame','Endgame Notes','Inherent Robot Ability','Did Well','Vulnerabilities']
pitColumns = ['Team Name', 'Drive Train Type', 'Wheel Type', 'Number of Wheels', 'Robot Speed', 'Robot Weight', 'Has Auto?', 'No Auto', 'Switch Auto', 'Scale Auto', 'Auto Line', '# Switch Auto', '# Scale Auto', 'Auto Comments', '# Switch Teleop', '# Scale Teleop', '# Exchange Teleop', 'Climb?', 'Support Other?', 'Teleop Comments']
def compute_sort(allTeams, valueList, teamList, string_col):
#     print allTeams
    for team in allTeams:
#         print team
        try:
            value = teamDictofAverages[team][string_col]
        except KeyError:
#             print "KEYERROR"
            value = 0
#         print "VALUE", team, value
        if len(valueList) == 0:
            teamList.append(team)
            valueList.append(value)
        else:
            for index in range(0, len(valueList)):
                if value > valueList[index]:
                    teamList.insert(index, team)
                    valueList.insert(index, value)
                    break
                elif index == len(valueList)-1:
                    teamList.append(team)
                    valueList.append(value)
    print len(valueList), "VAL LIST", valueList
#     print "TEAM LIST", teamList
    return teamList, valueList
def is_int(theInt):
    try:
        int(theInt)
        return True
    except ValueError:
        return False
def unicode_working(value):
    try:
        str(value)
        return True
    except UnicodeEncodeError:
        return False
def in_sheet(row, column, sheet):
    try:
        sheet.cell(row, column)
        return True
    except IndexError:
        return False
def can_write(row, column, sheet, value):
    try:
        sheet.write(row, column, value)
        return True
    except KeyError:
        return False
for row in range(1, numberOfRowsMatch):
    if in_sheet(row, 3, matchsheet) is False:
        break
    else:
#         print row, allMatchRowTeamNumbers
        allMatchRowTeamNumbers.append(str(int((matchsheet.cell(row, 3).value))))
for row in range(1, numberOfRowsPit):
    if in_sheet(row, 2, pitsheet) is False:
        break
    else:
#         print row, allMatchRowTeamNumbers
        allMatchRowTeamNumbers.append(str(int((pitsheet.cell(row, 2).value))))
for row in range(0, len(allMatchRowTeamNumbers)):
    if allMatchRowTeamNumbers[row] not in allTeams:
        allTeams.append(allMatchRowTeamNumbers[row])
# print allMatchRowTeamNumbers
# print allTeams
newBook = xlwt.Workbook()
for team in allTeams:
    allTeams[allTeams.index(team)] = int(allTeams[allTeams.index(team)])
allTeams.sort()
for team in allTeams:
    allTeams[allTeams.index(team)] = str(allTeams[allTeams.index(team)])
# print allTeams
for team in allTeams:
    teamReportSheet = newBook.add_sheet(team)
    teamToAverage = {}
    #MATCH SCOUTING REPORT
    teamReportSheet.write(0,0,'Robot Report for Team '+team)
    for col in matchColumns:
        teamReportSheet.write(1,matchColumns.index(col),col)
#         print col
    rowNumber = 1
    for row in range(1, numberOfRowsMatch):
        if in_sheet(row, 3, matchsheet) == False:
            break
        colNumber = -1
#         print team
#         print matchsheet.cell(row, 3).value, team
        if int(matchsheet.cell(row, 3).value) == int(team):
            rowNumber += 1
#             print rowNumber, colNumber
            matchCols = [2,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18]
            for match in matchCols:
                colNumber +=1
                print matchsheet.cell(row, match).value, row, match
                if match == 2:
                    teamReportSheet.write(rowNumber,colNumber,int(matchsheet.cell(row, match).value))
                    teamToAverage[matchColumns[colNumber]] = int(matchsheet.cell(row, match).value)
                else:
                    if unicode_working(matchsheet.cell(row, match).value):
                        teamReportSheet.write(rowNumber,colNumber,int(str(matchsheet.cell(row, match).value)[0]) if str(matchsheet.cell(row, match).value) != "" and is_int(str(matchsheet.cell(row, match).value)[0]) and is_int(str(matchsheet.cell(row, match).value)[1]) == False else str(matchsheet.cell(row,match).value))
                        teamToAverage[matchColumns[colNumber]] = int(str(matchsheet.cell(row, match).value)[0]) if str(matchsheet.cell(row, match).value) != "" and is_int(str(matchsheet.cell(row, match).value)[0]) and is_int(str(matchsheet.cell(row, match).value)[1]) == False else str(matchsheet.cell(row,match).value)
#     print teamToAverage
    allMatchRowData[team] = teamToAverage
    #PIT SCOUTING REPORT
    rowNumber += 2
    for col in pitColumns:
        teamReportSheet.write(rowNumber,pitColumns.index(col),col)
    
    for row in range(1, numberOfRowsPit):
        if in_sheet(row, 2, pitsheet) == False:
            break
        colNumber = -1
#         print team
#         print pitsheet.cell(row, 2).value, team
        if int(pitsheet.cell(row, 2).value) == int(team):
            rowNumber += 1
#             print rowNumber
            for pit in range(3, 23):
                colNumber += 1
#                 print pitsheet.cell(row, pit).value
                teamReportSheet.write(rowNumber, colNumber, int(str(pitsheet.cell(row, pit).value)[0]) if str(pitsheet.cell(row, pit).value) != "" and is_int(str(pitsheet.cell(row, pit).value)[0]) and is_int(str(pitsheet.cell(row, pit).value)[1]) == False else str(pitsheet.cell(row, pit).value))

teamDictofAverages = {}
#average the columns
for team in allTeams:
    listOfAverages = {}
#     print allMatchRowData[team].keys()
    for col in allMatchRowData[team].keys():
        adding = 0
#         print col
        for add in range(0, len(allMatchRowData[team])):
            toAdd = allMatchRowData[team][col]
            if is_int(toAdd) and col != "Match Number":
                adding += int(toAdd)
        adding /= len(allMatchRowData[team])
        listOfAverages[col] = adding
    teamDictofAverages[team] = listOfAverages
print teamDictofAverages, "teamDictofAverages"
teamListSortedByInherentAbility, valueListSortedByInherentAbility = compute_sort(allTeams, [], [], "Inherent Robot Ability")
teamListSortedByAutoPerformance, valueListSortedByAutoPerformance = compute_sort(allTeams, [], [], "Autonomous Performance")
teamListSortedByManeuverability, valueListSortedByManeuverability = compute_sort(allTeams, [], [], "Maneuverability")
teamListSortedBySwitchDelivery, valueListSortedBySwitchDelivery = compute_sort(allTeams, [], [], "Switch Cube Delivery")
teamListSortedByScaleDelivery, valueListSortedByScaleDelivery = compute_sort(allTeams, [], [], "Scale Cube Delivery")
teamListSortedByExchangeDelivery, valueListSortedByExchangeDelivery = compute_sort(allTeams, [], [], "Exchange Cube Delivery")
teamListSortedByEndgame, valueListSortedByEndgame = compute_sort(allTeams, [], [], "Endgame")

# print len(teamListSortedByEndgame), teamListSortedByEndgame

bestSheet = newBook.add_sheet("Team Rankings")

#Best Overall
bestOverallList = ['Inherent Robot Ability','Scale Cube Delivery','Autonomous Performance','Switch Cube Delivery','Endgame','Exchange Cube Delivery','Maneuverability','Driver Skill']
bestSheet.write(0, 0, "BEST OVERALL")
bestSheet.write(1, 0, "OUR RANK")
bestSheet.write(1, 1, "Team Number")
for bestOverallHeaders in range(0, len(bestOverallList)):
    bestSheet.write(1, bestOverallHeaders+2, bestOverallList[bestOverallHeaders])

for bestOverall in range(0, len(teamListSortedByInherentAbility)):
    print bestOverall, "BO"
    teamNumber = teamListSortedByInherentAbility[bestOverall]
#     bestSheet.write(bestOverall+2, 0, bestOverall+1)
    bestSheet.write(bestOverall+2, 1, teamListSortedByInherentAbility[bestOverall])
    for write in bestOverallList:
        if valueListSortedByInherentAbility[teamListSortedByInherentAbility.index(teamNumber)] != 0:
            can_write(bestOverall+2, bestOverallList.index(write)+2, bestSheet, teamDictofAverages[teamListSortedByInherentAbility[bestOverall]][write])

#Inherent Robot
#Scale Delivery
#Auto Performance
#Switch Delivery
#Endgame
#Exchange Delivery


#THIS MUST BE LAST
fileErrorNumber = 1
try:
    newBook.save('Individual_Scouting_Reports.xlsx')
except IOError:
    while True:
        try:
            newBook.save('Individual_Scouting_Reports_' + str(fileErrorNumber)+'.xlsx')
            break
        except IOError:
            fileErrorNumber += 1
            continue
