import pandas as pd
import gspread
import json
import re

weeklyCATCampaignMonthSlashYear = input("input campaign month slash year mm/yyyy: ")

def getNextMonth(monthSlashYear):
    month = monthSlashYear[0:2]
    year = monthSlashYear[3:]

    if int(month) + 1 == 13:
        month = 1
        year = int(year) + 1
    else:
        month = int(month) + 1
    
    if len(str(month)) == 1:
        month = f"0{month}"

    nextMonthSlashYear = f"{month}/{year}"
    return nextMonthSlashYear

def weeklyCAT(campaignMonthSlashYear = weeklyCATCampaignMonthSlashYear):
    credsList = [
        r"json files\credentials\credentials1.json",
        r"json files\credentials\credentials2.json",
        r"json files\credentials\credentials3.json"
    ]

    testCreds = r"CAT's automatization project\json files\test.json"
    gc = gspread.service_account(filename = testCreds)

    with open(r"CAT's automatization project\json files\block links.json", 'r+') as f:
        gSheetsIDsOfBlockFiles = json.load(f)
    
    for key in gSheetsIDsOfBlockFiles:
        sh = gc.open_by_key(gSheetsIDsOfBlockFiles[key])
        
        indexOfMonthOperational = sh.worksheet('Operational tab').get_values("1:1")[0].index(campaignMonthSlashYear)
        listOfWorksheetsMonth = sh.worksheet('Operational tab').col_values(indexOfMonthOperational + 1)
        monthWeeksDict = {}

        for campaign in listOfWorksheetsMonth[1: listOfWorksheetsMonth.index("Blank")]:
            ws = sh.worksheet(campaign)
            print(campaign)

            listOfCampaignMonthes = []
            dColValues = ws.get_values('D:D')
            for value in dColValues:
                regexResult = re.search("\d{2}/\d{4}", value[0])
                if regexResult == None:
                    continue
                listOfCampaignMonthes.append(value)

            if campaignMonthSlashYear == listOfCampaignMonthes[-1][0]:
                rowOfMonthPlusOne = dColValues.index([campaignMonthSlashYear]) + 2
                monthWeeks = ws.get_values(f'A{rowOfMonthPlusOne}:A250')
            else:
                rowOfMonthPlusOne = dColValues.index([campaignMonthSlashYear]) + 2
                indexOfMonthNext = listOfCampaignMonthes.index([campaignMonthSlashYear]) + 1
                NextMonth = listOfCampaignMonthes[indexOfMonthNext][0]
                rowOfNextMonth = dColValues.index([NextMonth])
                monthWeeks = ws.get_values(f'A{rowOfMonthPlusOne}:A{rowOfNextMonth}')

            for week in monthWeeks:
                if week[0] not in monthWeeksDict.keys():
                    monthWeeksDict[week[0]] = [campaign]

                if campaign not in monthWeeksDict[week[0]]:
                    monthWeeksDict[week[0]].append(campaign)

        if len(monthWeeksDict.keys()) == 0:
            continue

        monthWeeksDictSorted = {}
        monthWeeksDictSortedKeys = list(monthWeeksDict.keys())
        monthWeeksDictSortedKeys.sort()

        for key in monthWeeksDictSortedKeys:
            monthWeeksDictSorted[key] = monthWeeksDict[key]
        print(monthWeeksDictSorted)
        
        weeklyWS = sh.worksheet('Weekly stats')
        pandasCopyA = pd.DataFrame(weeklyWS.get_values("A2:A"))
        pandasCopyD = pd.DataFrame(weeklyWS.get_values("D2:D"))
        pandasCopy = pd.concat([pandasCopyA, pandasCopyD], axis=1)
        pandasCopy.columns = ["month", "dateCID"]
        print(pandasCopy)

        thisMonthStart = pandasCopy.where(pandasCopy==campaignMonthSlashYear).first_valid_index()
        thisMonthEnd = pandasCopy.where(pandasCopy==campaignMonthSlashYear).last_valid_index()

        if thisMonthStart == None:
            print(pandasCopy.last_valid_index())
        print(pandasCopy.where(pandasCopy=='49').last_valid_index())
        print(thisMonthStart, thisMonthEnd)
        # ws.update("D640", 49)
        break


weeklyCAT()