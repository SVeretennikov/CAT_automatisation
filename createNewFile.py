import gspread
import json

def createNewFile(type, uniqueName, creds):
    gc = gspread.service_account(filename=creds)

    if type == "block":
        nameTemplate = f'CAT\'s {uniqueName} Block launched campaigns'
        folderID = "1poZk2THvnJeuKGGGyQAHZrqYHPeHrSmi"
        template = "1Q3LBmsTXQa-XWzRNs8zen17LHOnKvyahRiq7xiJvUCU"
        jsonFilePath = r"CAT's automatization project\json files\block links.json"
    else:
        nameTemplate = f'CAT\'s {uniqueName} launched campaigns'
        folderID = "1gFGnV7sCJsWGJreMABUb3FGnAxiA_eVy"
        template = "1AgCC0tAnEfnZzOqqUm7oBtAk3-rf5Xh_1GGyNfA4E7U"
        jsonFilePath = r"CAT's automatization project\json files\CAT members.json"


    gc.copy(template, title=nameTemplate, copy_permissions=True, folder_id=folderID, copy_comments=False)
    newFileID = gc.open(nameTemplate).id
    gc.insert_permission(newFileID, "s.veretennikov@infuseua.com", perm_type="user", role="owner")
    with open(jsonFilePath, 'r+') as file:
        data = json.load(file)
        data[str(uniqueName)] = newFileID
        print(data)
        file.seek(0)
        json.dump(data, file, indent= 4)
        updatedTypeDict = json.load(file)

    return updatedTypeDict
    