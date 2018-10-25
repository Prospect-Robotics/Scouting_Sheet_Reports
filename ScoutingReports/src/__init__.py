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
allMatchRowTeamNumbers = []
allMatchRowData = []
allTeams = []
matchColumns = ['Match Number','Alliance Color','Robot Starting Position','Autonomous Performance','Autonomous Notes','Switch Cube Delivery','Scale Cube Delivery','Exchange Cube Delivery','Maneuverability','Driver Skill','Teleoperation Notes','Endgame','Endgame Notes','Inherent Robot Ability','Did Well','Vulnerabilities']
pitColumns = ['Team Name', 'Drive Train Type', 'Wheel Type', 'Number of Wheels', 'Robot Speed', 'Robot Weight', 'Has Auto?', 'No Auto', 'Switch Auto', 'Scale Auto', 'Auto Line', '# Switch Auto', '# Scale Auto', 'Auto Comments', '# Switch Teleop', '# Scale Teleop', '# Exchange Teleop', 'Climb?', 'Support Other?', 'Teleop Comments']
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
        print row, allMatchRowTeamNumbers
        allMatchRowTeamNumbers.append(str(int((matchsheet.cell(row, 3).value))))
for row in range(1, numberOfRowsPit):
    if in_sheet(row, 2, pitsheet) == False:
        break
    else:
        print row, allMatchRowTeamNumbers
        allMatchRowTeamNumbers.append(str(int((pitsheet.cell(row, 2).value))))
for row in range(0, len(allMatchRowTeamNumbers)):
    if allMatchRowTeamNumbers[row] not in allTeams:
        allTeams.append(allMatchRowTeamNumbers[row])
print allMatchRowTeamNumbers
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
    for col in matchColumns:
        teamReportSheet.write(1,matchColumns.index(col),col)
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
    for col in pitColumns:
        teamReportSheet.write(rowNumber,pitColumns.index(col),col)
    
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
bestSheet = newBook.add_sheet("Team Rankings")

#Best Overall
bestSheet.write(0, 0, "BEST OVERALL")
bestSheet.write(1, 0, "OUR RANK")
bestSheet.write(1, 1, "Team Number")
bestSheet.write(1, 2, "Inherent Robot Ability (Match avg)")
bestSheet.write(1, 3, "Maneuverability (Match avg)")
bestSheet.write(1, 4, "Driver Skill")
bestSheet.write(1, 5, "Endgame")
bestSheet.write(1, 6, "Teleoperation Notes")
bestSheet.write(1, 7, "Did Well")
bestSheet.write(1, 8, "Vulnerabilities")

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
