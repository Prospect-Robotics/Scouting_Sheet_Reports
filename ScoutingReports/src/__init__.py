import xlrd
import xlwt
ourbook = raw_input("Enter the filename of the Scouting Spreadsheet: ")
workbook = xlrd.open_workbook(ourbook+'.xlsx')
#oursheet = raw_input("Enter the sheet name within the spreadsheet: ")
matchsheet = workbook.sheet_by_name('MATCH SCOUTING')
pitsheet = workbook.sheet_by_name('PIT SCOUTING')
#ourRows = input("Enter the number of filled rows in the spreadsheet: ")
numberOfRowsMatch = input("Enter the number of filled rows in Match sheet: ")
numberOfRowsPit = input("Enter the number of filled rows in Pit sheet: ")
allRowTeamNumbers = []
allTeams = []
def is_int(theInt):
    try:
        int(theInt)
        return True
    except ValueError:
        return False
def in_sheet(row, column, sheet):
    try:
        sheet.cell(row, column)
        return True
    except IndexError:
        return False
for row in range(1, numberOfRowsMatch):
    if in_sheet(row, 3, matchsheet) == False:
        break
    else:
        print row, allRowTeamNumbers
        allRowTeamNumbers.append(str(int((matchsheet.cell(row, 3).value))))
for row in range(1, numberOfRowsPit):
    if in_sheet(row, 2, pitsheet) == False:
        break
    else:
        print row, allRowTeamNumbers
        allRowTeamNumbers.append(str(int((pitsheet.cell(row, 2).value))))
for row in range(0, len(allRowTeamNumbers)):
    if allRowTeamNumbers[row] not in allTeams:
        allTeams.append(allRowTeamNumbers[row])
print allRowTeamNumbers
print allTeams
newBook = xlwt.Workbook()
for team in allTeams:
    allTeams[allTeams.index(team)] = int(allTeams[allTeams.index(team)])
allTeams.sort()
for team in allTeams:
    allTeams[allTeams.index(team)] = str(allTeams[allTeams.index(team)])
print allTeams
for team in allTeams:
    teamReportSheet = newBook.add_sheet(team)

    #MATCH SCOUTING REPORT
    teamReportSheet.write(0,0,'Robot Report for Team '+team)
    teamReportSheet.write(1,0,'Match Number')
    teamReportSheet.write(1,1,'Alliance Color')#'Blue' or 'Red'
    teamReportSheet.write(1,2,'Robot Starting Position')#[11:-1] is 'Left' or 'Center' or 'Right'
    teamReportSheet.write(1,3,'Autonomous Performance')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,4,'Autonomous Notes')#Long answer
    teamReportSheet.write(1,5,'Switch Cube Delivery')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,6,'Scale Cube Delivery')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,7,'Exchange Cube Delivery')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,8,'Maneuverability')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,9,'Driver Skill')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,10,'Teleoperation Notes')#Long answer
    teamReportSheet.write(1,11,'Endgame')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,12,'Endgame Notes')#Long answer
    teamReportSheet.write(1,13,'Inherent Robot Ability')#[0] is 1, 2, 3, or 4
    teamReportSheet.write(1,14,'Did Well')#Long answer
    teamReportSheet.write(1,15,'Vulnerabilities')#Long answer
    rowNumber = 1
    for row in range(1, numberOfRowsMatch):
        if in_sheet(row, 3, matchsheet) == False:
            break
        colNumber = -1
        print team
        print matchsheet.cell(row, 3).value, team
        if int(matchsheet.cell(row, 3).value) == int(team):
            rowNumber += 1
            print rowNumber, colNumber
            matchCols = [2,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
            for match in matchCols:
                colNumber +=1
                print matchsheet.cell(row, match).value
                if match == 2:
                    teamReportSheet.write(rowNumber,colNumber,int(matchsheet.cell(row, match).value))
                else:
                    teamReportSheet.write(rowNumber,colNumber,int(str(matchsheet.cell(row, match).value)[0]) if str(matchsheet.cell(row, match).value) != "" and is_int(str(matchsheet.cell(row, match).value)[0]) and is_int(str(matchsheet.cell(row, match).value)[1]) == False else str(matchsheet.cell(row,match).value))

    #PIT SCOUTING REPORT
    rowNumber += 2
    teamReportSheet.write(rowNumber, 0, 'Team Name')
    teamReportSheet.write(rowNumber, 1, 'Drive Train Type')
    teamReportSheet.write(rowNumber, 2, 'Wheel Type')
    teamReportSheet.write(rowNumber, 3, 'Number of Wheels')
    teamReportSheet.write(rowNumber, 4, 'Robot Speed')
    teamReportSheet.write(rowNumber, 5, 'Robot Weight')
    teamReportSheet.write(rowNumber, 6, 'Has Auto?')
    teamReportSheet.write(rowNumber, 7, 'No Auto')
    teamReportSheet.write(rowNumber, 8, 'Switch Auto')
    teamReportSheet.write(rowNumber, 9, 'Scale Auto')
    teamReportSheet.write(rowNumber, 10, 'Auto Line')
    teamReportSheet.write(rowNumber, 11, '# Switch Auto')
    teamReportSheet.write(rowNumber, 12, '# Scale Auto')
    teamReportSheet.write(rowNumber, 13, 'Auto Comments')
    teamReportSheet.write(rowNumber, 14, '# Switch Teleop')
    teamReportSheet.write(rowNumber, 15, '# Scale Teleop')
    teamReportSheet.write(rowNumber, 16, '# Exchange Teleop')
    teamReportSheet.write(rowNumber, 17, 'Climb?')
    teamReportSheet.write(rowNumber, 18, 'Support Other?')
    teamReportSheet.write(rowNumber, 19, 'Teleop Comments')
    
    for row in range(1, numberOfRowsPit):
        if in_sheet(row, 2, pitsheet) == False:
            break
        colNumber = -1
        print team
        print pitsheet.cell(row, 2).value, team
        if int(pitsheet.cell(row, 2).value) == int(team):
            rowNumber += 1
            print rowNumber
            for pit in range(3, 23):
                colNumber += 1
                print pitsheet.cell(row, pit).value
                teamReportSheet.write(rowNumber, colNumber, int(str(pitsheet.cell(row, pit).value)[0]) if str(pitsheet.cell(row, pit).value) != "" and is_int(str(pitsheet.cell(row, pit).value)[0]) and is_int(str(pitsheet.cell(row, pit).value)[1]) == False else str(pitsheet.cell(row, pit).value))

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
