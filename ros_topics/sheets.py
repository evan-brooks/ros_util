from datetime import datetime
import json
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1xfrqQuTaNfXBFwPcpCcDYYpPDvO8MgLEWMq9AY7vbqY'
CREDS_FILE = '../google_api_creds.json'
GOOGLE_API_TOKEN = '../token.json'

def google_auth() -> Credentials:
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(GOOGLE_API_TOKEN):
        creds = Credentials.from_authorized_user_file(GOOGLE_API_TOKEN, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(GOOGLE_API_TOKEN, 'w') as token:
            token.write(creds.to_json())

    return creds

def create_new_sheet(service) -> str:
    # Generate the current date and time as the sheet name
    sheet_name = datetime.now().strftime("%Y-%m-%d %I:%M%p")
    
    # Create a new sheet with the generated name
    new_sheet_request = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": sheet_name
                    }
                }
            }
        ]
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=new_sheet_request
    ).execute()

    new_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
    
    return (sheet_name, new_sheet_id)


def set_column_widths_requests(sheet_id: str):
    column_widths=[310, 310, 75, 75]

    width_requests = []
    for i, width in enumerate(column_widths):
        width_request = {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": {
                    "pixelSize": width
                },
                "fields": "pixelSize"
            }
        }
        width_requests.append(width_request)
        
    return width_requests

def freeze_top_row_request(sheet_id: str):
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {
                    "frozenRowCount": 1  # Freeze 1 row
                }
            },
            "fields": "gridProperties.frozenRowCount"
        }
    }

def top_row_bold_request(sheet_id: str):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold": True
                    }
                }
            },
            "fields": "userEnteredFormat.textFormat.bold"
        }
    }

def set_colum_decimal_format_request(sheet_id, column_index, decimal_count):
    """Set the number of decimals for a specific column."""
    
    # The number format pattern based on the desired decimal count
    pattern = f"0.{ '0' * decimal_count }"
    
    # The batch update request body
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startColumnIndex": column_index,
                "endColumnIndex": column_index + 1,
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {
                        "type": "NUMBER",
                        "pattern": pattern
                    }
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    }

def format_sheet(service, sheet_id: str):
    requests = []

    requests.extend(set_column_widths_requests(sheet_id))
    requests.append(freeze_top_row_request(sheet_id))
    requests.append(top_row_bold_request(sheet_id))
    requests.append(set_colum_decimal_format_request(sheet_id, column_index=3, decimal_count=3))
    requests.append(set_colum_decimal_format_request(sheet_id, column_index=4, decimal_count=1))

    print(json.dumps(requests, indent=2))

    service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()

def update_spreadsheet(topics_data: list[tuple[str, str, float, float]]):
    creds = google_auth()

    # Initialize the Sheets API service
    service = build('sheets', 'v4', credentials=creds)

    (sheet_name, sheet_id) = create_new_sheet(service)    
    format_sheet(service, sheet_id)

    header = ['topic', 'type', 'mbps', 'hz']
    data_lists = [header]
    
    # Convert data from list of tuples to list of lists
    data_lists.extend([list(item) for item in topics_data])
    
    # Create the request body
    body = {
        'values': data_lists
    }
    
    # Make the update request
    range_name = f"'{sheet_name}'!A:D"
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()
    
    print(f"{result.get('updatedCells')} cells updated.")
    return sheet_id