from __future__ import print_function

import os.path
import textstat

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a spreadsheet.
SPREADSHEET_ID = '1XSOzWHLWIeZgntDPPP85vKgQAnVxNPqhlPbs8OjRdLA'
DATA_RANGE = 'Cleaned!A2:E'
NEW_RANGE = 'Cleaned!D2:D'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
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

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=DATA_RANGE).execute()
        values = result.get('values')

        if not values:
            print('No data found.')
            return

        scores = []

        for row in values:
            scores.append(textstat.text_standard(row[0], float_output=True))
        
        body = {
            'range': NEW_RANGE,
            'values': [scores],
            'majorDimension': "Columns"
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, 
            range=NEW_RANGE,
            body=body,
            valueInputOption="RAW"
            
            ).execute()
        print(f"{result.get('updatedCells')} cells updated.")

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()