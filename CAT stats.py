import time
import requests
import json
import pandas as pd
import os
import base64
import sqlite3
import shutil
from datetime import datetime, timedelta
from win32 import win32crypt
from Crypto.Cipher import AES
import gspread
from dateutil.parser import parse

def get_chrome_datetime(chromedate):
    if chromedate != 86400000000 and chromedate:
        try:
            return datetime(1601, 1, 1) + timedelta(microseconds=chromedate)
        except Exception as e:
            print(f"Error: {e}, chromedate: {chromedate}")
            return chromedate
    else:
        return ""

def get_encryption_key():
    local_state_path = os.path.join(os.environ["USERPROFILE"],
                                    "AppData", "Local", "Google", "Chrome",
                                    "User Data", "Local State")
    with open(local_state_path, "r", encoding="utf-8") as f:
        local_state = f.read()
        local_state = json.loads(local_state)

    key = base64.b64decode(local_state["os_crypt"]["encrypted_key"])
    key = key[5:]
    return win32crypt.CryptUnprotectData(key, None, None, None, 0)[1]

def decrypt_data(data, key):
    try:
        iv = data[3:15]
        data = data[15:]
        cipher = AES.new(key, AES.MODE_GCM, iv)
        return cipher.decrypt(data)[:-16].decode()
    except:
        try:
            return str(win32crypt.CryptUnprotectData(data, None, None, None, 0)[1])
        except:
            return ""

def extract_cookies_from_google_crome():
    db_path = os.path.join(os.environ["USERPROFILE"], "AppData", "Local",
                            "Google", "Chrome", "User Data", "Default", "Network", "Cookies")
    filename = "Cookies.db"
    if not os.path.isfile(filename):
        shutil.copyfile(db_path, filename)
    db = sqlite3.connect(filename)
    db.text_factory = lambda b: b.decode(errors="ignore")
    cursor = db.cursor()
    cursor.execute("""
    SELECT host_key, name, value, encrypted_value 
    FROM cookies
    WHERE host_key like 'lv.infusemedia.com'""")
    key = get_encryption_key()
    for host_key, name, value, encrypted_value in cursor.fetchall():
        if not value:
            decrypted_value = decrypt_data(encrypted_value, key)
        else:
            decrypted_value = value
        cookieNameValueDict[name] = decrypted_value
        cursor.execute("""
        UPDATE cookies SET value = ?, has_expires = 1, expires_utc = 99999999999999999, is_persistent = 1, is_secure = 0
        WHERE host_key = ?
        AND name = ?""", (decrypted_value, host_key, name))
    db.commit()
    db.close()

def enter_list_data(rowNumber):
    ws.batch_update([
        {'range': enterValuesMapping['date'] + str(rowNumber), 'values': [[campaignDaySlashMonthSlashYear]]},
        {'range': enterValuesMapping['link'] + str(rowNumber), 'values': [["https://lv.infusemedia.com/list/"+str(CATRow['id'])]]},
        {'range': enterValuesMapping['list name/month'] + str(rowNumber), 'values': [[CATRow['manual_name']]]},
        {'range': enterValuesMapping['sent'] + str(rowNumber), 'values': [[CATRow['stats']['sent']]]},
        {'range': enterValuesMapping['accepted ov'] + str(rowNumber), 'values': [[CATRow['stats']['accepted_ov']]]},
        {'range': enterValuesMapping['accepted pv'] + str(rowNumber), 'values': [[CATRow['stats']['accepted_pv']]]},
        {'range': enterValuesMapping['unsucessful pv'] + str(rowNumber), 'values': [[CATRow['stats']['unsuccessfull_pv']]]},
        {'range': enterValuesMapping['q3'] + str(rowNumber), 'values': [[CATRow['stats']['other_q3']]]},
        {'range': enterValuesMapping['title green'] + str(rowNumber), 'values': [[CATRow['stats']['title_reject_green']]]},
        {'range': enterValuesMapping['title yellow'] + str(rowNumber), 'values': [[CATRow['stats']['title_reject_yellow']]]},
        {'range': enterValuesMapping['country green'] + str(rowNumber), 'values': [[CATRow['stats']['country_reject_green']]]},
        {'range': enterValuesMapping['country yellow'] + str(rowNumber), 'values': [[CATRow['stats']['country_reject_yellow']]]},
        {'range': enterValuesMapping['industry green'] + str(rowNumber), 'values': [[CATRow['stats']['industry_reject_green']]]},
        {'range': enterValuesMapping['industry yellow'] + str(rowNumber), 'values': [[CATRow['stats']['industry_reject_yellow']]]},
        {'range': enterValuesMapping['emp size green'] + str(rowNumber), 'values': [[CATRow['stats']['employees_reject_green']]]},
        {'range': enterValuesMapping['emp size yellow'] + str(rowNumber), 'values': [[CATRow['stats']['employees_reject_yellow']]]},
        {'range': enterValuesMapping['revenue green'] + str(rowNumber), 'values': [[CATRow['stats']['revenue_reject_green']]]},
        {'range': enterValuesMapping['revenue yellow'] + str(rowNumber), 'values': [[CATRow['stats']['revenue_reject_yellow']]]},
        {'range': enterValuesMapping['nac sup contact'] + str(rowNumber), 'values': [[CATRow['stats']['contact_nac_sup_reject']]]},
        {'range': enterValuesMapping['nac sup company'] + str(rowNumber), 'values': [[CATRow['stats']['company_nac_sup_reject']]]},
        {'range': enterValuesMapping['nwc ov'] + str(rowNumber), 'values': [[CATRow['stats']['nwc_ov']]]},
        {'range': enterValuesMapping['bad data'] + str(rowNumber), 'values': [[CATRow['stats']['bad_data']]]},
        {'range': enterValuesMapping['out of business'] + str(rowNumber), 'values': [[CATRow['stats']['out_of_business']]]},
        {'range': enterValuesMapping['nwc pv'] + str(rowNumber), 'values': [[CATRow['stats']['nwc_pv']]]},
        {'range': enterValuesMapping['na proof'] + str(rowNumber), 'values': [[CATRow['stats']['prooflink_na']]]},
        {'range': enterValuesMapping['na other contact']    + str(rowNumber), 'values': [[CATRow['stats']['contact_other_na']]]},
        {'range': enterValuesMapping['na other company'] + str(rowNumber), 'values': [[CATRow['stats']['company_other_na']]]},
        {'range': enterValuesMapping['na duplicate'] + str(rowNumber), 'values': [[CATRow['stats']['duplicate_na']]]},
        {'range': enterValuesMapping['na backup verified'] + str(rowNumber), 'values': [[CATRow['stats']['backup_verified']]]},
        {'range': enterValuesMapping['q1'] + str(rowNumber), 'values': [[CATRow['stats']['qtitle']]]},
        {'range': enterValuesMapping['q2'] + str(rowNumber), 'values': [[CATRow['stats']['qcompany']]]},
        ], value_input_option = 'user_entered')

def is_date(string, fuzzy=False):
    try: 
        parse(string, fuzzy=fuzzy)
        return True
    except ValueError:
        return False

def deploy_metadata_to_new_campaign_ws():
    sh.worksheet(CATRow['campaign_cid'][0:7]).update('D3', str(campaignMonthSlashYear), value_input_option = 'user_entered')
    sh.worksheet(CATRow['campaign_cid'][0:7]).update('CL1', sh.worksheet(CATRow['campaign_cid'][0:7]).id, value_input_option = 'user_entered')
    sh.worksheet(CATRow['campaign_cid'][0:7]).update('CM1', sh.worksheet(CATRow['campaign_cid'][0:7]).title, value_input_option = 'user_entered')

def find_where_to_put_month(campaignYearMonthes):
    indexesOfMonthThatAreMoreThanListMonth = [i for i, e in enumerate(campaignYearMonthes) if datetime.strptime(e, '%m/%Y').month > datetime.strptime(campaignMonthSlashYear, '%m/%Y').month]
    if indexesOfMonthThatAreMoreThanListMonth == [] or campaignYearMonthes == []:
        requiredIndex = 1
        if indexesOfYearsThatAreMoreThanListYear == []:
            return listOfListsOfDatesAndLVListNames.index(listOfListsOfDatesAndLVListNames[-1]) + 4
        else:
            rowOfNextYear = listOfListsOfDatesAndLVListNames.index([listOfDates[indexesOfYearsThatAreMoreThanListYear[0]]]) + 3
            return rowOfNextYear
    else:
        requiredIndex = indexesOfMonthThatAreMoreThanListMonth[0]
        requiredMonth = campaignYearMonthes[requiredIndex]
        return listOfListsOfDatesAndLVListNames.index([requiredMonth]) + 3

def insert_new_line(row):
    newRowInputList = gc.open_by_key(gSheetsIDsOfBlockFiles[f'{block}']).worksheet('Campaign template').get_values('6:6', value_render_option='formula')[0]
    ws.insert_row([input.replace('6', str(row)) for input in newRowInputList], row, value_input_option = 'user_entered')

def insert_new_month_and_list():
    insert_new_line(find_where_to_put_month(listOfCampaignYearMonthes))
    insert_new_line(find_where_to_put_month(listOfCampaignYearMonthes))
    ws.update('D'+str(find_where_to_put_month(listOfCampaignYearMonthes)), campaignMonthSlashYear, value_input_option = 'user_entered')
    enter_list_data(find_where_to_put_month(listOfCampaignYearMonthes)+1)

def deploy_new_totals(indexWhereToDeploy, sh, campaignMonthSlashYear):
    sh.worksheet('Totals template').duplicate(insert_sheet_index=indexWhereToDeploy, new_sheet_name=f'Totals ({campaignMonthSlashYear})')
    sh.worksheet(f'Totals ({campaignMonthSlashYear})').update('B2', str(campaignMonthSlashYear), value_input_option = 'user_entered')
    sh.worksheet(f'Totals ({campaignMonthSlashYear})').update('CJ1', sh.worksheet(f'Totals ({campaignMonthSlashYear})').id, value_input_option = 'user_entered')

def sorting_funk(cid):
    return cid[0][2:7]

def totals():
    print('Totals input started')
    listWithCredsForEachLVList = [credsList[x % 3] for x in range(len(gSheetsIDsOfBlockFiles))]
    credsMapping = 0
    for blockName, blockLink in gSheetsIDsOfBlockFiles.items():
        gc = gspread.service_account(filename=listWithCredsForEachLVList[credsMapping])
        credsMapping += 1
        sh = gc.open_by_key(blockLink)
        listOfWorksheets = [worksheet.title for worksheet in sh.worksheets()]
        weeklyStatsIndex = sh.worksheet('Weekly stats').index
        listOfTotals = listOfWorksheets[1:weeklyStatsIndex]
        listOfTotalsForAllStats = [totals[8:-1] for totals in listOfTotals if totals[8:-1] in setOfCATCampaignMonthesSlashYears]
        dictOfTotalsInBlockFiles[blockName] = listOfTotalsForAllStats
        for campaignMonthSlashYear in setOfCATCampaignMonthesSlashYears:
            if datetime.strptime(campaignMonthSlashYear, '%m/%Y').month in range(1, 10):
                comparableCATCampaignMonthSlashYear = str(datetime.strptime(campaignMonthSlashYear, '%m/%Y').year) + '0' + str(datetime.strptime(campaignMonthSlashYear, '%m/%Y').month)
            else:
                comparableCATCampaignMonthSlashYear = str(datetime.strptime(campaignMonthSlashYear, '%m/%Y').year) + str(datetime.strptime(campaignMonthSlashYear, '%m/%Y').month)
            
            listOfCampaignWorksheets = listOfWorksheets[(weeklyStatsIndex+1):len(listOfWorksheets)-2]

            if f'Totals ({campaignMonthSlashYear})' not in listOfWorksheets:
                listOfCampaignsThatHaveListsInCATCampaignMonth = []
                for CID in listOfCampaignWorksheets:
                    listOfValuesOnCID = sh.worksheet(CID).get_values('D3:D200')
                    # hereyo
                    if [campaignMonthSlashYear] in listOfValuesOnCID:
                        listOfCampaignsThatHaveListsInCATCampaignMonth.append(CID)
                        print(CID)
                        # campaignMonthSlashYearIndex = listOfValuesOnCID.index([campaignMonthSlashYear])
                        # if campaignMonthSlashYear[0:2] == '12':
                        #     campaignMonthSlashYearPlusOne = "01/" + str(int(campaignMonthSlashYear[3:]) + 1)
                        # elif len(str(int(campaignMonthSlashYear[0:2]) + 1)) == 1:
                        #     campaignMonthSlashYearPlusOne = "0" + str(int(campaignMonthSlashYear[0:2]) + 1) + campaignMonthSlashYear[2:]
                        # else:
                        #     campaignMonthSlashYearPlusOne = str(int(campaignMonthSlashYear[0:2]) + 1) + campaignMonthSlashYear[2:]
                    time.sleep(2)
                        
                if listOfCampaignsThatHaveListsInCATCampaignMonth != []:
                    if listOfTotals == []:
                        deploy_new_totals(weeklyStatsIndex, sh, campaignMonthSlashYear)
                    else:
                        listOfComparableTotalsDates = []
                        for totals in listOfTotals:  
                            if datetime.strptime(totals[8:15], '%m/%Y').month in range(1, 10):
                                listOfComparableTotalsDates.append(str(datetime.strptime(totals[8:15], '%m/%Y').year) + '0' + str(datetime.strptime(totals[8:15], '%m/%Y').month))
                            else:
                                listOfComparableTotalsDates.append(str(datetime.strptime(totals[8:15], '%m/%Y').year) + str(datetime.strptime(totals[8:15], '%m/%Y').month))            
                        totalsLessThanCurrentTotals = [i for i in listOfComparableTotalsDates if i > comparableCATCampaignMonthSlashYear]
                        if totalsLessThanCurrentTotals != []:
                            deploy_new_totals(listOfWorksheets.index(listOfTotals[listOfComparableTotalsDates.index(totalsLessThanCurrentTotals[0])]), sh, campaignMonthSlashYear)
                        else:
                            deploy_new_totals(weeklyStatsIndex, sh, campaignMonthSlashYear)

                    monthCol = sh.worksheet('Operational tab').find(campaignMonthSlashYear, 1).address[0]
                    sh.worksheet('Operational tab').update(monthCol + '2:' + monthCol + str(len(listOfCampaignsThatHaveListsInCATCampaignMonth)+1), [listOfCampaignsThatHaveListsInCATCampaignMonth], major_dimension='COLUMNS')
                    time.sleep(20)
                    print(f'{blockName} Totals ({campaignMonthSlashYear}) is ready')
                else:
                    time.sleep(2)
                    print(f'{blockName} Totals ({campaignMonthSlashYear}) is ready') 
                    continue    
                        
            elif f'Totals ({campaignMonthSlashYear})' in listOfWorksheets:
                listOfCampaignsThatHaveListsInCATCampaignMonth = sh.worksheet(f'Totals ({campaignMonthSlashYear})').get_values('B3:B250')
                listOfCampaignsThatHaveListsInCATCampaignMonth = listOfCampaignsThatHaveListsInCATCampaignMonth[0:listOfCampaignsThatHaveListsInCATCampaignMonth.index(['Blank'])]

                listOfCampaignWorksheets = [campaignCID for campaignCID in listOfCampaignWorksheets if campaignCID not in listOfCampaignsThatHaveListsInCATCampaignMonth]

                for worksheet in listOfCampaignWorksheets:
                    print(worksheet)
                    if [worksheet] in listOfCampaignsThatHaveListsInCATCampaignMonth:
                        continue
                    elif [campaignMonthSlashYear] in sh.worksheet(worksheet).get_values('D3:D250'):
                        listOfCampaignsThatHaveListsInCATCampaignMonth.append([worksheet])
                    time.sleep(1)
                    
                listOfCampaignsThatHaveListsInCATCampaignMonth.sort(key=sorting_funk)
                listOfCampaignsThatHaveListsInCATCampaignMonth = [finalCID[0] for finalCID in listOfCampaignsThatHaveListsInCATCampaignMonth]

                if listOfCampaignsThatHaveListsInCATCampaignMonth == []:
                    time.sleep(2)
                    print(f'{blockName} Totals ({campaignMonthSlashYear}) is ready') 
                    continue
                else:
                    monthCol = sh.worksheet('Operational tab').find(campaignMonthSlashYear, 1).address[0]
                    sh.worksheet('Operational tab').update(monthCol + '2:' + monthCol + str(len(listOfCampaignsThatHaveListsInCATCampaignMonth)+1), [listOfCampaignsThatHaveListsInCATCampaignMonth], major_dimension='COLUMNS')
                
                time.sleep(20)
                print(f'{blockName} Totals ({campaignMonthSlashYear}) is ready')

def responsible_cats():
    print('Adding responsible CATs to Totals')
    dictBlocksMonthesCats = {}
    for blockName, blockLink in gSheetsIDsOfBlockFiles.items():
        for campaignMonthSlashYear in setOfCATCampaignMonthesSlashYears:
            sh = gc.open_by_key(blockLink)
            existingWorksheets = sh.worksheets()
            listOfWorksheets = [worksheet.title for worksheet in existingWorksheets]
            listCATsThisBlockThisMonth = []
            dictCATsThisMonth = {}
            if f'Totals ({campaignMonthSlashYear})' in listOfWorksheets:
                totalsWorksheet = sh.worksheet(f'Totals ({campaignMonthSlashYear})')
                listOfTotalsCIDs = [cid for cid in totalsWorksheet.get_values('B3:B1000') if cid != ['Blank']]
                totalsWorksheetMonthmmslashyyyy = totalsWorksheet.get('B2')
                if datetime.strptime(totalsWorksheetMonthmmslashyyyy[0][0], '%m/%Y').month in range(1, 10):
                    totalsWorksheetMonthyyyymm = str(datetime.strptime(totalsWorksheetMonthmmslashyyyy[0][0], '%m/%Y').year) + '0' + str(datetime.strptime(totalsWorksheetMonthmmslashyyyy[0][0], '%m/%Y').month)
                else:
                    totalsWorksheetMonthyyyymm = str(datetime.strptime(totalsWorksheetMonthmmslashyyyy[0][0], '%m/%Y').year) + str(datetime.strptime(totalsWorksheetMonthmmslashyyyy[0][0], '%m/%Y').month)
                rowForCATInput = 3
                for totalsCID in listOfTotalsCIDs:
                    if len(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]) == 1:
                        totalsWorksheet.update('CI'+str(rowForCATInput), dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[0])
                        listCATsThisBlockThisMonth.append(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[0])
                    else: 
                        listOfStartDates = list(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['start date'])
                        for i, startDate in enumerate(listOfStartDates):
                            if datetime.strptime(startDate, '%m/%d/%Y').month in range(1,10):
                                listOfStartDates[i] = str(datetime.strptime(startDate, '%m/%d/%Y').year) + '0' + str(datetime.strptime(startDate, '%m/%d/%Y').month)
                            else:
                                listOfStartDates[i] = str(datetime.strptime(startDate, '%m/%d/%Y').year) + str(datetime.strptime(startDate, '%m/%d/%Y').month)
                        if totalsWorksheetMonthyyyymm in listOfStartDates:
                            listCATsThisBlockThisMonth.append(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[listOfStartDates.index(totalsWorksheetMonthyyyymm)])
                            totalsWorksheet.update('CI'+str(rowForCATInput), dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[listOfStartDates.index(totalsWorksheetMonthyyyymm)])
                        else:
                            listOfMonthesThatMoreThanTotalsNumber = [month for month in listOfStartDates if totalsWorksheetMonthyyyymm > month]
                            listCATsThisBlockThisMonth.append(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[listOfStartDates.index(listOfMonthesThatMoreThanTotalsNumber[-1])])
                            totalsWorksheet.update('CI'+str(rowForCATInput), dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(totalsCID[0])]['Responsible CAT'].iloc[listOfStartDates.index(listOfMonthesThatMoreThanTotalsNumber[-1])])

                    print(f'{totalsCID[0]}')
                    rowForCATInput += 1
                    time.sleep(2)
                dictCATsThisMonth[campaignMonthSlashYear] = listCATsThisBlockThisMonth
                print(f'{blockName} Totals ({campaignMonthSlashYear}) is ready')
            else:
                continue
        dictBlocksMonthesCats[blockName] = dictCATsThisMonth
    return dictBlocksMonthesCats

def all_cats_stats():
    print('Entering all stats')
    sh = gc.open_by_key('1Yl88ezZ5BdJ3xzE1PQD3kaJx7RLeSuNr-VWvhDDXbxc')
    for campaignMonthSlashYear in setOfCATCampaignMonthesSlashYears:
        existingWorksheets = sh.worksheets()
        listOfWorksheets = [worksheet.title for worksheet in existingWorksheets]
        listOfAllTotals = listOfWorksheets[2:-1]
        if f'Totals ({campaignMonthSlashYear})' not in listOfAllTotals:
            compareableCampaignMonthSlashYear = str(campaignMonthSlashYear[3:]) + str(campaignMonthSlashYear[0:2])
            totalsCompareableMOnthes = [str(totalsdate[11:-1])+str(totalsdate[8:10]) for totalsdate in listOfAllTotals if str(totalsdate[11:-1])+str(totalsdate[8:10]) > compareableCampaignMonthSlashYear]
            if totalsCompareableMOnthes == []:
                sh.worksheet('Totals template').duplicate(insert_sheet_index=len(listOfWorksheets)-1, new_sheet_name=f'Totals ({campaignMonthSlashYear})')
                sh.worksheet(f'Totals ({campaignMonthSlashYear})').update('B2', str(campaignMonthSlashYear), value_input_option = 'user_entered')
            else:
                sh.worksheet('Totals template').duplicate(insert_sheet_index=listOfWorksheets.index(f'Totals ({totalsCompareableMOnthes[0][-2:]}/{totalsCompareableMOnthes[0][0:4]})'), new_sheet_name=f'Totals ({campaignMonthSlashYear})')
                sh.worksheet(f'Totals ({campaignMonthSlashYear})').update('B2', str(campaignMonthSlashYear), value_input_option = 'user_entered')
        ws = sh.worksheet('Operational tab')
        dictOfTotalsInBlockFileskeysThatHaveCampaignMonthSlashYear = [k + ' Block' for k, v in dictOfTotalsInBlockFiles.items() if campaignMonthSlashYear in v]
        monthCol = ws.find(campaignMonthSlashYear, 1).address[0]
        ws.update(monthCol + '2:' + monthCol + str(len(dictOfTotalsInBlockFileskeysThatHaveCampaignMonthSlashYear)+1), [dictOfTotalsInBlockFileskeysThatHaveCampaignMonthSlashYear], major_dimension='COLUMNS')

def dupe_individual_block_template(index, block, campaignMonthSlashYear, sh):
    sh.worksheet('Block template').duplicate(insert_sheet_index=index, new_sheet_name=f'{block} Block')
    sh.worksheet(f'{block} Block').update('D3', str(campaignMonthSlashYear), value_input_option = 'user_entered')
    sh.worksheet(f'{block} Block').update('CL1', sh.worksheet(f'{block} Block').id, value_input_option = 'user_entered')
    sh.worksheet(f'{block} Block').update('CM1', f'{block} Block', value_input_option = 'user_entered')

def enter_individual_files(dictBlocksMonthesCats):
    listWithCredsForEachLVList = [credsList[x % 3] for x in range(len(dictCATMembers))]
    credsMapping = 0
    print(dictBlocksMonthesCats)
    for CAT, CATLink in dictCATMembers.items():
        for campaignMonthSlashYear in setOfCATCampaignMonthesSlashYears:
            print(CAT)
            gc = gspread.service_account(filename=listWithCredsForEachLVList[credsMapping])
            credsMapping += 1
            sh = gc.open_by_key(CATLink)
            listOfBlocksCATHasThisMonth = []
            for block, monthSlashYear in dictBlocksMonthesCats.items():
                if dictBlocksMonthesCats[block] == {}:
                    continue
                listOfCATInDictBlocksMonthesCats = [CATInDictBlocksMonthesCats for CATInDictBlocksMonthesCats in dictBlocksMonthesCats[block][campaignMonthSlashYear] if CATInDictBlocksMonthesCats == CAT]
                existingWorksheets = sh.worksheets()
                listOfWorksheets = [worksheet.title for worksheet in existingWorksheets]
                weeklyStatsIndex = sh.worksheet('Weekly stats').index
                if CAT not in dictBlocksMonthesCats[block][campaignMonthSlashYear]:
                    continue
                elif f'{block} Block' not in listOfWorksheets:
                    listOfBlocks = listOfWorksheets[weeklyStatsIndex+1:-2]
                    listOfBlocksThatMoreThanCurrent = [block for block in listOfBlocks if block > f'{block} Block']
                    if listOfBlocks == []:
                        dupe_individual_block_template(weeklyStatsIndex+1, block, campaignMonthSlashYear, sh)
                    elif listOfBlocksThatMoreThanCurrent == []:
                        dupe_individual_block_template(listOfWorksheets.index(listOfWorksheets[-2]), block, campaignMonthSlashYear, sh)
                    else:
                        dupe_individual_block_template(listOfWorksheets.index(listOfBlocksThatMoreThanCurrent[0]), block, campaignMonthSlashYear, sh)
                    time.sleep(5)
                elif f'{block} Block' in listOfWorksheets:
                    listOfDatesAndCIDs = sh.worksheet(f'{block} Block').get_values('D3:D250')
                    if [campaignMonthSlashYear] not in listOfDatesAndCIDs:
                        listOfEnteredDates = [dateCID for dateCID in listOfDatesAndCIDs if dateCID[0][0:2] != block]
                        dateRowList = sh.worksheet('Block template').get_values('3:3', value_render_option='formula')[0]
                        filterImportrangeRowList = sh.worksheet('Block template').get_values('4:4', value_render_option='formula')[0]
                        additionalCampaignRowList = sh.worksheet('Block template').get_values('5:5', value_render_option='formula')[0]
                        monthesThatAreMoreThanCampaignMonthSlashYear = [date for date in listOfEnteredDates if str(date[0][3:7])+str(date[0][0:2]) > str(campaignMonthSlashYear[3:7])+str(campaignMonthSlashYear[0:2])]
                        if monthesThatAreMoreThanCampaignMonthSlashYear == []:
                            indexOfRequiredCell = listOfDatesAndCIDs.index(listOfDatesAndCIDs[-1]) + 4
                        else:
                            indexOfRequiredCell = listOfDatesAndCIDs.index(monthesThatAreMoreThanCampaignMonthSlashYear[0]) + 3
                        rowsRange = range(indexOfRequiredCell, indexOfRequiredCell + 1 + len(listOfCATInDictBlocksMonthesCats))
                        listOfListsOfRows = []
                        print(f'{block} Block')
                        for uniqueRow in rowsRange:
                            additionalCampaignRowListCopy = []
                            uniqueRowMinusOne = uniqueRow - 1
                            dateFormula = f'=IFS(ISDATE(D{uniqueRow})=True, TEXT(D{uniqueRow}, "mm/yyyy"), D{uniqueRow}="", "", ISDATE(D{uniqueRow})=False, DATEVALUE(B{uniqueRowMinusOne}))'
                            if len(listOfListsOfRows) == 0:
                                for i, e in enumerate(dateRowList):
                                    if i == 1:                                               
                                        dateRowList[i] = dateFormula
                                    elif i == 2:
                                        dateRowList[i] = dateRowList[i].replace('D3', f'D{uniqueRow}')
                                    elif i == 3:
                                        dateRowList[i] = campaignMonthSlashYear
                                    else:
                                        dateRowList[i] = dateRowList[i].replace('3', f'{uniqueRow}')
                                listOfListsOfRows.append(dateRowList)
                            elif len(listOfListsOfRows) == 1:
                                for i, e in enumerate(filterImportrangeRowList):
                                    if i == 1:                                               
                                        filterImportrangeRowList[i] = dateFormula
                                    elif i == 2:
                                        filterImportrangeRowList[i] = filterImportrangeRowList[i].replace('D4', f'D{uniqueRow}').replace('B4', f'B{uniqueRow}')
                                    elif i == 3:
                                        filterImportrangeRowList[i] = filterImportrangeRowList[i].replace('D3', f'D{uniqueRowMinusOne}')
                                    else:
                                        filterImportrangeRowList[i] = filterImportrangeRowList[i].replace('4', f'{uniqueRow}')
                                        filterImportrangeRowList[i] = filterImportrangeRowList[i].replace('D3', f'D{uniqueRowMinusOne}')
                                listOfListsOfRows.append(filterImportrangeRowList)
                            else:
                                for i, e in enumerate(additionalCampaignRowList):
                                    if i == 1:
                                        additionalCampaignRowListCopy.append(dateFormula)
                                    elif i == 2:
                                        additionalCampaignRowListCopy.append(additionalCampaignRowList[i].replace('D5', f'D{uniqueRow}').replace('B5', f'B{uniqueRow}'))
                                    elif additionalCampaignRowList[i] == '':
                                        additionalCampaignRowListCopy.append('')
                                    else:
                                        additionalCampaignRowListCopy.append(additionalCampaignRowList[i].replace('5', f'{uniqueRow}'))
                                listOfListsOfRows.append(additionalCampaignRowListCopy)
                        sh.worksheet(f'{block} Block').insert_rows(listOfListsOfRows, indexOfRequiredCell, value_input_option = 'USER_ENTERED')
                        time.sleep(5)
                if CAT in monthSlashYear[campaignMonthSlashYear]:
                    listOfBlocksCATHasThisMonth.append(f'{block} Block')

            if listOfBlocksCATHasThisMonth != []:
                existingWorksheets = sh.worksheets()
                listOfWorksheets = [worksheet.title for worksheet in existingWorksheets]
                weeklyStatsIndex = sh.worksheet('Weekly stats').index
                listOfTotals = listOfWorksheets[1:weeklyStatsIndex]
                if f'Totals ({campaignMonthSlashYear})' in listOfTotals:
                    pass
                elif listOfTotals == []:
                    deploy_new_totals(1, sh, campaignMonthSlashYear)
                else:
                    listOfTotalsMoreThanCurrentMonth = [totals for totals in listOfTotals if totals[11:-1] + totals[8:10] > campaignMonthSlashYear[3:7] + campaignMonthSlashYear[0:2]]
                    if listOfTotalsMoreThanCurrentMonth == []:
                        deploy_new_totals(weeklyStatsIndex, sh, campaignMonthSlashYear)
                    else:
                        deploy_new_totals(sh.worksheet(listOfTotalsMoreThanCurrentMonth[0]).index, sh, campaignMonthSlashYear)
                listOfBlocksCATHasThisMonth.sort()
                monthCol = sh.worksheet('Operational tab').find(campaignMonthSlashYear, 1).address[0]
                sh.worksheet('Operational tab').update(monthCol + '2:' + monthCol + str(len(listOfBlocksCATHasThisMonth)+1), [listOfBlocksCATHasThisMonth], major_dimension='COLUMNS')

startDate = input('Enter start date yyyy-mm-dd format:')
endDate = input('Enter end date yyyy-mm-dd format:')

credsList = [
    r"json files\credentials\credentials1.json",
    r"json files\credentials\credentials2.json",
    r"json files\credentials\credentials3.json"
]

with open(r"json files\block links.json", 'r+') as f:
    gSheetsIDsOfBlockFiles = json.load(f)

with open(r"json files\CAT members.json", 'r+') as f:
    dictCATMembers = json.load(f)

enterValuesMapping = {
    'date': 'B',
    'link': 'C',
    'list name/month': 'D',
    'sent': 'K',
    'accepted ov': 'M',
    'accepted pv': 'O',
    'unsucessful pv': 'Q',
    'q3': 'S',
    'title green': 'AE',
    'title yellow': 'AG',
    'country green': 'AK',
    'country yellow': 'AM',
    'industry green': 'AQ',
    'industry yellow': 'AS',
    'emp size green': 'AW',
    'emp size yellow': 'AY',
    'revenue green': 'BC',
    'revenue yellow': 'BE',
    'nac sup contact': 'BG',
    'nac sup company': 'BI',
    'nwc ov': 'BM',
    'bad data': 'BO',
    'out of business': 'BQ',
    'nwc pv': 'BS',
    'na proof': 'BU',
    'na other contact': 'BW',
    'na other company': 'BY',
    'na duplicate': 'CA',
    'na backup verified': 'CE',
    'q1': 'CG',
    'q2': 'CI'
    }

cookieNameValueDict = {}

extract_cookies_from_google_crome()

statisticsForCampaignsRequestURL = "https://lv.infusemedia.com/api/stats/campaign"
querystring = {"date_mode":"custom","date_custom_from":startDate,"date_custom_to":endDate}
craftingCookieString = 'remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d=' + cookieNameValueDict['remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d'] + '; io=' + cookieNameValueDict['io'] + '; AWSALB=' + cookieNameValueDict['AWSALB'] + '; AWSALBCORS=' + cookieNameValueDict['AWSALBCORS'] + '; XSRF-TOKEN=' + cookieNameValueDict['XSRF-TOKEN'] + '; lvprod_session=' + cookieNameValueDict['lvprod_session']

payload = ""
headers = {
    "cookie": craftingCookieString,
    "authority": "lv.infusemedia.com",
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9",
    "guid": "8de5cd6f-c1d7-542b-b66c-db592e578865",
    "referer": "https://lv.infusemedia.com/stats/campaign",
    "sec-ch-ua": "^\^.Not/A",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "^\^Windows^^",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36",
    "x-csrf-token": "jh8u4P0BCfYKQh13qJNoyh7wpeuJkNWnn7f1FJLs",
    "x-requested-with": "XMLHttpRequest",
    "x-xsrf-token": cookieNameValueDict['XSRF-TOKEN']
}

response = requests.request("GET", statisticsForCampaignsRequestURL, data=payload, headers=headers, params=querystring)
jsonAllStats = json.loads(response.text)

gc = gspread.service_account(filename=r"json files\credentials\credentials_for_campaign_monitoring.json")
sh = gc.open_by_key('178-z2ljURQa5WX9pVwvcIQhwxxHQgEXGo6X-PJyDvEM')
ws = sh.worksheet('CATs Monitor')

startDateList = ws.get_values('A2:A1000')
responsibleCATList = ws.get_values('C2:C1000')
cidList = ws.get_values('E2:E1000')

for i, startDate in enumerate(startDateList):
    startDateList[i] = str(startDate[0])
for i, e in enumerate(responsibleCATList):
    if e == [''] or e == ['without CAT']:
        responsibleCATList[i] = None
    else:
        responsibleCATList[i] = str(e[0])
for i, cidFromList in enumerate(cidList):
    cidList[i] = str(cidFromList[0])

dataFromCampaignMonitoring = pd.DataFrame({
    'start date' : [date for date in startDateList][0:len([cat for cat in responsibleCATList])], 
    'Responsible CAT' : [cat for cat in responsibleCATList],
    'CID' : [cid[0:7] for cid in cidList][0:len([cat for cat in responsibleCATList])]
})
dataFromCampaignMonitoring.dropna(how='any', inplace=True)

CATRows = [row for row in jsonAllStats['data']['rows'] if row['campaign_cid'][0:7] in list(dataFromCampaignMonitoring['CID'])]
listWithCredsForEachLVList = [credsList[x % 3] for x in range(len(CATRows))]
credsMapping = 0
dpoSMLKeywords = ['DPO_', 'SML_', 'Return', 'return', 'investigation']
setOfCATCampaignDates = {CATRow['created_at'][0:10] for CATRow in CATRows}
setOfCATCampaignMonthesSlashYears = set()

for CATCampaignDate in setOfCATCampaignDates:
    yyyydmmddd = datetime.strptime(CATCampaignDate, '%Y-%m-%d')
    if yyyydmmddd.month in range(1, 10):
        campaignMonth = '0' + str(yyyydmmddd.month)
    else:
        campaignMonth = str(yyyydmmddd.month)
    campaignMonthSlashYear = campaignMonth + '/' + str(yyyydmmddd.year)
    setOfCATCampaignMonthesSlashYears.add(campaignMonthSlashYear)

dictOfTotalsInBlockFiles = {}

amountOfLists = len(CATRows)
print(amountOfLists)

for i, CATRow in enumerate(CATRows, 1):
    print(f"{i}/{amountOfLists} {CATRow['campaign_cid'][0:7]}")

    gc = gspread.service_account(filename=listWithCredsForEachLVList[credsMapping])
    credsMapping += 1

    yyyydmmddd = datetime.strptime(CATRow['created_at'][0:10], '%Y-%m-%d')
    if yyyydmmddd.month in range(1, 10):
        campaignMonth = '0' + str(yyyydmmddd.month)
    else:
        campaignMonth = str(yyyydmmddd.month)
    if yyyydmmddd.day in range(1, 10):
        campaignDay = '0' + str(yyyydmmddd.day)
    else:
        campaignDay = str(yyyydmmddd.day)
    campaignMonthSlashYear = campaignMonth + '/' + str(yyyydmmddd.year)
    campaignDaySlashMonthSlashYear = campaignMonth + '/' + campaignDay + '/' + str(yyyydmmddd.year)
    campaignYyyyMmDd = str(yyyydmmddd.year) + campaignMonth + campaignDay
    dateOfCampaignFirstStart = list(dataFromCampaignMonitoring[dataFromCampaignMonitoring['CID'].str.contains(CATRow['campaign_cid'][0:7])]['start date'])[0]

    dateOfCampaignFirstStartyyyydmmddd = datetime.strptime(dateOfCampaignFirstStart, '%m/%d/%Y')
    if dateOfCampaignFirstStartyyyydmmddd.month in range(1, 10):
        campaignFirstStartMonth = '0' + str(dateOfCampaignFirstStartyyyydmmddd.month)
    else:
        campaignFirstStartMonth = str(dateOfCampaignFirstStartyyyydmmddd.month)
    if dateOfCampaignFirstStartyyyydmmddd.day in range(1, 10):
        campaignFirstStartDay = '0' + str(dateOfCampaignFirstStartyyyydmmddd.day)
    else:
        campaignFirstStartDay = str(dateOfCampaignFirstStartyyyydmmddd.day)
    campaignFirstStartYyyyMmDd = str(dateOfCampaignFirstStartyyyydmmddd.year) + campaignFirstStartMonth + campaignFirstStartDay
    
    print(CATRow['manual_name'])
    block = CATRow['campaign_cid'][0:2]
    if block not in list(gSheetsIDsOfBlockFiles.keys()):
        raise Exception("Not finished feature yet")
        # gc.copy(gSheetsIDsOfBlockFiles['template'], title=f'CAT\'s {block} Block launched campaigns', copy_permissions=True, folder_id='10PR2tI3T5hlnOdi-dXsGIXXt7pfomqfh', copy_comments=False)
        # with open(r"CAT's automatization project\json files\block links.json", 'r+') as file:
        #     data = json.load(file)
        #     data[str(block)] =  gc.open(f'CAT\'s {block} Block launched campaigns').id
        #     print(data)
        #     file.seek(0)
        #     json.dump(data, file, indent= 4)
        # gSheetsIDsOfBlockFiles[f'{block}'] =  gc.open(f'CAT\'s {block} Block launched campaigns').id
    sh = gc.open_by_key(gSheetsIDsOfBlockFiles[f'{block}'])

    weeklyStatsIndex = sh.worksheet('Weekly stats').index
    existingWorksheets = sh.worksheets()    
    listOfWorksheets = [worksheet.title for worksheet in existingWorksheets]

    if CATRow['campaign_cid'][0:7] not in listOfWorksheets:
        listOfCIDs = listOfWorksheets[weeklyStatsIndex+1:-2]
        listOfCIDsWhichHaveBiggerNumber = [i for i in listOfCIDs if i[2:7] > CATRow['campaign_cid'][2:7]]
        if any(dpoSMLKeyword in CATRow['manual_name'] for dpoSMLKeyword in dpoSMLKeywords) == True:
            print('DPO/SML/Return list')
            time.sleep(10)
            continue
        elif (listOfCIDsWhichHaveBiggerNumber == ['Weekly stats']) or listOfCIDsWhichHaveBiggerNumber == []:
            sh.worksheet('Campaign template').duplicate(insert_sheet_index=listOfWorksheets.index(listOfWorksheets[-2]), new_sheet_name=CATRow['campaign_cid'][0:7])
            deploy_metadata_to_new_campaign_ws()
        else:
            sh.worksheet('Campaign template').duplicate(insert_sheet_index=listOfWorksheets.index(listOfCIDsWhichHaveBiggerNumber[0]), new_sheet_name=CATRow['campaign_cid'][0:7])
            deploy_metadata_to_new_campaign_ws()
    
    ws = sh.worksheet(CATRow['campaign_cid'][0:7])
    listOfListsOfDatesAndLVListNames = ws.get_values('D3:D250')
    listOfListsOfPlatformLinksAndTotalLinks = ws.get_values('C3:C250')
    
    accepted = CATRow['stats']['sent'] + CATRow['stats']['accepted_ov'] + CATRow['stats']['accepted_pv'] + CATRow['stats']['unsuccessfull_pv'] - CATRow['stats']['other_q3'] - CATRow['stats']['nwc_pv']
    totalAccepted = accepted + CATRow['stats']['other_q3']
    totalTitle = CATRow['stats']['title_reject_green'] + CATRow['stats']['title_reject_yellow']
    totalCountry = CATRow['stats']['country_reject_green'] + CATRow['stats']['country_reject_yellow']
    totalIndustry = CATRow['stats']['industry_reject_green'] + CATRow['stats']['industry_reject_yellow']
    totalEmpSize = CATRow['stats']['employees_reject_green'] + CATRow['stats']['employees_reject_yellow']
    totalRevenue = CATRow['stats']['revenue_reject_green'] + CATRow['stats']['revenue_reject_yellow']
    totalCriterion = totalTitle + totalCountry + totalIndustry + totalEmpSize + totalRevenue + CATRow['stats']['contact_nac_sup_reject'] + CATRow['stats']['company_nac_sup_reject']
    totalBadData = CATRow['stats']['nwc_ov'] + CATRow['stats']['bad_data'] + CATRow['stats']['out_of_business'] + CATRow['stats']['nwc_pv'] + CATRow['stats']['prooflink_na'] + CATRow['stats']['contact_other_na'] + CATRow['stats']['company_other_na'] + CATRow['stats']['duplicate_na']
    totalAllRejects = totalCriterion + totalBadData
    totalZombie = CATRow['stats']['backup_verified'] + CATRow['stats']['qtitle'] + CATRow['stats']['qcompany']
    totalChecked = totalAccepted + totalAllRejects + totalZombie

    listOfReturnsToIsDateCheck = []
    listOfDates = []
    if ws.get_values('B2:B250') == []:
        enter_list_data(len(list(filter(None, ws.col_values(4))))+1)
    elif ([CATRow['manual_name']] in listOfListsOfDatesAndLVListNames) and (["https://lv.infusemedia.com/list/"+str(CATRow['id'])] in listOfListsOfPlatformLinksAndTotalLinks):
        enteredListRowNumber = listOfListsOfDatesAndLVListNames.index([CATRow['manual_name']]) + 3
        currentValueOfTotalChecked = ws.get_values(f"E{enteredListRowNumber}")
        if [[str(totalChecked)]] == currentValueOfTotalChecked:
            print('list is already entered')
            time.sleep(10)
            continue
        else:
            enter_list_data(enteredListRowNumber)
    elif campaignFirstStartYyyyMmDd > campaignYyyyMmDd:
        print('campaign is not CAT campaign yet')
        time.sleep(10)
        continue
    elif any(dpoSMLKeyword in CATRow['manual_name'] for dpoSMLKeyword in dpoSMLKeywords) == True:
        print('DPO/SML/Return list')
        time.sleep(10)
        continue
    elif [campaignMonthSlashYear] not in listOfListsOfDatesAndLVListNames:
        [listOfReturnsToIsDateCheck.append(is_date(list[0])) for list in listOfListsOfDatesAndLVListNames]
        listOfTrueIndexes = [i for i, e in enumerate(listOfReturnsToIsDateCheck) if e == True]
        [listOfDates.append(listOfListsOfDatesAndLVListNames[trueIndex][0]) for trueIndex in listOfTrueIndexes]
        indexesOfYearsThatAreLessThanListYear = [i for i, e in enumerate(listOfDates) if datetime.strptime(e, '%m/%Y').year < datetime.strptime(campaignMonthSlashYear, '%m/%Y').year]
        indexesOfYearsThatAreMoreThanListYear = [i for i, e in enumerate(listOfDates) if datetime.strptime(e, '%m/%Y').year > datetime.strptime(campaignMonthSlashYear, '%m/%Y').year]
        if (indexesOfYearsThatAreLessThanListYear == []) and (indexesOfYearsThatAreMoreThanListYear == []):
            listOfCampaignYearMonthes = listOfDates
            insert_new_month_and_list()
        elif indexesOfYearsThatAreLessThanListYear == []:
            listOfCampaignYearMonthes = listOfDates[0:indexesOfYearsThatAreMoreThanListYear[0]]
            insert_new_month_and_list()
        elif indexesOfYearsThatAreMoreThanListYear == []:
            listOfCampaignYearMonthes = listOfDates[indexesOfYearsThatAreLessThanListYear[-1]+1:len(listOfDates)]
            insert_new_month_and_list()
        else:
            listOfCampaignYearMonthes = listOfDates[indexesOfYearsThatAreLessThanListYear[-1]+1:indexesOfYearsThatAreMoreThanListYear[0]]
            insert_new_month_and_list()
    else:
        indexOfListMonth = listOfListsOfDatesAndLVListNames.index([campaignMonthSlashYear]) + 3
        if [''] not in ws.get_values(f'B{indexOfListMonth + 1}:B250'):
            datesOfListMonth = ws.get_values(f'B{indexOfListMonth + 1}:B250')
        else:
            datesOfListMonth = ws.get_values(f'B{indexOfListMonth + 1}:B250')[0:ws.get_values(f'B{indexOfListMonth + 1}:B250').index([''])]
        indexesOfDaysThatIsMoreThanCampaignDay = [i for i, e in enumerate(datesOfListMonth) if datetime.strptime(e[0], '%m/%d/%Y').day > datetime.strptime(campaignDaySlashMonthSlashYear, '%m/%d/%Y').day]
        if indexesOfDaysThatIsMoreThanCampaignDay == []:
            if datesOfListMonth[-1][0] == ws.get_values(f'B{indexOfListMonth + 1}:B250')[-1][0]:
                enter_list_data(indexOfListMonth + len(datesOfListMonth)+1)
            else:
                insert_new_line(indexOfListMonth + len(datesOfListMonth)+1)
                enter_list_data(indexOfListMonth + len(datesOfListMonth)+1)
        else:
            insert_new_line(indexOfListMonth + len(datesOfListMonth[0:indexesOfDaysThatIsMoreThanCampaignDay[0]]) + 1)
            enter_list_data(indexOfListMonth + len(datesOfListMonth[0:indexesOfDaysThatIsMoreThanCampaignDay[0]]) + 1)
    print('list is ready to go')
    time.sleep(21)
    
totals()
dictBlocksMonthesCats = responsible_cats()
all_cats_stats()
enter_individual_files(dictBlocksMonthesCats)

print('Fin.')