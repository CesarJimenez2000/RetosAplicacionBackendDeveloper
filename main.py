from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID = '1DRD97TAw2WIuTCG0Nh6BW-aVvDKAgY1wvJb38V-3vU8'
OUTPUT_SPREADSHEET_ID = '1TqqIXRJELqcGd9YAQCOqB8S_S7ZIxWhepoSE9a-BVNY'
SAMPLE_RANGE_NAME = 'Reto1'
def numberToBase(n, b):
    if n == 0:
        return [0]
    digits = []
    while n:
        digits.append(int(n % b))
        n //= b
    return digits[::-1]

def getRangeByIndex(sheet,row,column):
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    newColumnIndexes = numberToBase(column,26)
    newColumn = ''
    for i in newColumnIndexes:
        newColumn += LETTERS[i]
    range = sheet + '!' + newColumn + str(row+1)
    return range

def writeHeader(service,spreadSheetID,name,row,column,size):
    body = {'values': [[name]]}
    service.spreadsheets().values().update(
        spreadsheetId=spreadSheetID, range=getRangeByIndex('Hoja 1',row,column),
        valueInputOption='USER_ENTERED',
        body=body).execute()
    requests = [
        {
            "mergeCells": {
                "range": {
                    "sheetId": '0',
                    "startRowIndex": row,
                    "endRowIndex": row+1,
                    "startColumnIndex": column,
                    "endColumnIndex": column+size
                },
                "mergeType": "MERGE_ALL",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": '0',
                    "startRowIndex": row,
                    "endRowIndex": row+1,
                    "startColumnIndex": column,
                    "endColumnIndex":column+size
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment)"
            }
        }
    ]
    body = {
        'requests': requests
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadSheetID,body=body).execute()
    return None

#this function formats the data and outputs it to the target spreadsheet:
#values, is the data retreived from the original spreadsheets
#index. is the index where the variable columns start

def formatOutput(service,spreadSheetID,values,index):
    headers = values[0]
    classes = []
    numVariable = len(headers)-index
    for i in range(numVariable):
        classes.append([])
    dataDict = dict()
    print(values)
    for x in range(1, len(values)):
        row = values[x]
        print(row)
        key = (row[0], row[1])
        if key in dataDict.keys():
            counter = index
            for i in dataDict[key]:
                i.append(row[counter])
                if not row[counter] in classes[counter-index]:
                    classes[counter - index].append(row[counter])
                counter += 1
        else:
            dataDict[key] = []
            for i in range(len(row)-index):
                dataDict[key].append([])
            counter = index
            for i in dataDict[key]:
                i.append(row[counter])
                if not row[counter] in classes[counter-index]:
                    classes[counter - index].append(row[counter])
                counter += 1
    indexCounter = index
    outputArray = []
    secondLine = []
    for i in range(index):
        secondLine.append(headers[i])
    for i in range(len(classes)):
        writeHeader(service,spreadSheetID,headers[index+i],0,indexCounter,len(classes[i]))
        indexCounter += len(classes[i])
        for variableType in classes[i]:
            secondLine.append(variableType)
    outputArray.append(secondLine)

    for key in dataDict:
        line = []
        for keyItem in key:
            line.append(keyItem)
        i = 0
        for classList in dataDict[key]:
            for variableType in classes[i]:
                if variableType in classList:
                    line.append('VERDADERO')
                else:
                    line.append('FALSO')
            i += 1
        outputArray.append(line)
    print(outputArray)
    body = {'values': outputArray}
    service.spreadsheets().values().update(
        spreadsheetId=spreadSheetID, range='Hoja 1!A2',
        valueInputOption='USER_ENTERED',
        body=body).execute()
    return dataDict


def main():
    creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range = SAMPLE_RANGE_NAME).execute()

        values = result.get('values', [])
        print(values)
        if not values:
            print('No data found on',SAMPLE_RANGE_NAME)
            return
        else:
            formatOutput(service,OUTPUT_SPREADSHEET_ID,values,2)
    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()