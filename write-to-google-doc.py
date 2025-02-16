from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pandas as pd
import os.path
import pickle

# Scopes required for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_google_auth():
    """
    Handle Google Sheets authentication
    Returns credentials object
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def get_excel_sheet_names(excel_file):
    """
    Get all sheet names from Excel file
    """
    try:
        xl = pd.ExcelFile(excel_file)
        return xl.sheet_names
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return []

def create_or_update_sheets(service, spreadsheet_id, sheet_names):
    """
    Create or update sheets in Google Spreadsheet to match Excel structure
    """
    try:
        # Get existing sheets
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing_sheets = {sheet['properties']['title'] for sheet in spreadsheet['sheets']}
        
        # Prepare batch update request
        batch_requests = []
        
        # Add new sheets if they don't exist
        for sheet_name in sheet_names:
            if sheet_name not in existing_sheets:
                batch_requests.append({
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                })
        
        # Execute batch update if needed
        if batch_requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': batch_requests}
            ).execute()
            
    except Exception as e:
        print(f"Error updating sheet structure: {str(e)}")
        raise

def update_google_sheet(spreadsheet_id, excel_file):
    """
    Update Google Spreadsheet with all sheets from Excel file
    """
    try:
        # Get authentication and build service
        creds = get_google_auth()
        service = build('sheets', 'v4', credentials=creds)
        
        # Get sheet names from Excel
        sheet_names = get_excel_sheet_names(excel_file)
        if not sheet_names:
            raise ValueError("No sheets found in Excel file")
        
        # Create or update sheets structure
        create_or_update_sheets(service, spreadsheet_id, sheet_names)
        
        # Clear and update each sheet
        for sheet_name in sheet_names:
            # Read data from Excel sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Convert DataFrame to list for Google Sheets
            data = [df.columns.values.tolist()] + df.values.tolist()
            
            # Clear existing data
            range_name = f"'{sheet_name}'!A1:ZZ"
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            # Update with new data
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_name}'!A1",
                valueInputOption='RAW',
                body={'values': data}
            ).execute()
            
            print(f"Successfully updated sheet: {sheet_name}")
        
        print("All sheets updated successfully")
        
    except Exception as e:
        print(f"Error updating Google Spreadsheet: {str(e)}")
        raise

def main():
    """
    Main function to run the script
    """
    try:
        
        # Get the Excel filename
        excel_filename = "your_local_spreadsheet.xlsx" # FIXME
        
        # Google Spreadsheet ID - Replace with your spreadsheet ID
        SPREADSHEET_ID = 'your_spreadsheet_id_here'  # FIXME
        
        # Update Google Spreadsheet
        update_google_sheet(SPREADSHEET_ID, excel_filename)
        
    except Exception as e:
        print(f"Error in main execution: {str(e)}")

if __name__ == "__main__":
    main()
    