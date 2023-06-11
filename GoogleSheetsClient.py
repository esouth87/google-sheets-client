import os
import pickle
import json
import google.auth
import google.auth.transport.requests
import googleapiclient.discovery
import googleapiclient.errors
from google_auth_oauthlib.flow import InstalledAppFlow

class GoogleSheetsClient:
    def __init__(self, credentials_file='credentials.json'):
        self.credentials_file = credentials_file
        self.credentials = self.get_credentials()

    def get_credentials(self):
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                credentials = pickle.load(token)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, ['https://www.googleapis.com/auth/spreadsheets'])
            credentials = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(credentials, token)
        return credentials

    def create_service(self):
        credentials = google.auth.default()[0] if self.credentials is None else self.credentials
        service = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials)
        return service

    def get_spreadsheet(self, spreadsheet_id):
        service = self.create_service()
        return service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    def write_cell(self, spreadsheet_id, sheet_name, row, col, value):
        service = self.create_service()
        sheet_range = f'{sheet_name}!{col}{row}'
        value_input_option = 'USER_ENTERED'
        body = {'values': [[value]]}
        return service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption=value_input_option,
            body=body
        ).execute()

    def get_next_empty_cell_in_column(self, spreadsheet_id, sheet_name, col):
        service = self.create_service()
        sheet_range = f'{sheet_name}!{col}:{col}'
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheet_range).execute()
        values = result.get('values', [])
        return len(values) + 1

    def get_next_empty_cell_in_row(self, spreadsheet_id, sheet_name, row):
        service = self.create_service()
        sheet_range = f'{sheet_name}!{row}'
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheet_range).execute()
        values = result.get('values', [])
        return len(values[0]) + 1

    def create_graph(self, spreadsheet_id, sheet_name, graph_title, data_range, graph_type='LINE'):
        service = self.create_service()
        chart_spec = {
            'title': graph_title,
            'basicChart': {
                'chartType': graph_type,
                'legendPosition': 'RIGHT_LEGEND',
                'axis': [
                    {'position': 'BOTTOM_AXIS', 'title': 'X-axis'},
                    {'position': 'LEFT_AXIS', 'title': 'Y-axis'}
                ],
                'domains': [{'domain': {'sourceRange': {'sources': [{'sheetId': sheet_name, 'startRowIndex': data_range['startRowIndex'], 'endRowIndex': data_range['endRowIndex'], 'startColumnIndex': data_range['startColumnIndex'], 'endColumnIndex': data_range['endColumnIndex']}]}}}],
                'series': [{'series': {'sourceRange': {'sources': [{'sheetId': sheet_name, 'startRowIndex': data_range['startRowIndex'], 'endRowIndex': data_range['endRowIndex'], 'startColumnIndex': data_range['startColumnIndex'] + 1, 'endColumnIndex': data_range['endColumnIndex']}]}}, 'targetAxis': 'LEFT_AXIS'}]
            }
        }
        requests = [
            {
                'addChart': {
                    'chart': {
                        'spec': chart_spec,
                        'position': {'newSheet': True}
                    }
                }
            }
        ]
        body = {'requests': requests}
        return service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

    def update_graph(self, spreadsheet_id, chart_id, data_range):
        service = self.create_service()
        chart_range = f'{data_range["sheet_name"]}!{data_range["start_col"]}:{data_range["end_col"]}'
        requests = [
            {
                'updateChartSpec': {
                    'chartId': chart_id,
                    'spec': {
                        'basicChart': {
                            'domains': [{'domain': {'sourceRange': {'sources': [{'sheetId': data_range['sheet_name'], 'startRowIndex': data_range['start_row_index'], 'endRowIndex': data_range['end_row_index'], 'startColumnIndex': data_range['start_col_index'], 'endColumnIndex': data_range['end_col_index']}]}}}],
                            'series': [{'series': {'sourceRange': {'sources': [{'sheetId': data_range['sheet_name'], 'startRowIndex': data_range['start_row_index'], 'endRowIndex': data_range['end_row_index'], 'startColumnIndex': data_range['start_col_index'] + 1, 'endColumnIndex': data_range['end_col_index']}]}}, 'targetAxis': 'LEFT_AXIS'}]
                        }
                    }
                }
            }
        ]
        body = {'requests': requests}
        return service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

    def create_sheet(self, spreadsheet_id, sheet_title):
        service = self.create_service()
        requests = [
            {
                'addSheet': {
                    'properties': {
                        'title': sheet_title
                    }
                }
            }
        ]
        body = {'requests': requests}
        return service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

    def watch_changes(self, spreadsheet_id, callback_url, sheet_name, row_or_col):
        service = self.create_service()
        request = {
            'type': 'LISTEN',
            'address': callback_url,
            'ranges': [f'{sheet_name}!{row_or_col}:{row_or_col}']
        }
        try:
            return service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=False).execute()
        except googleapiclient.errors.HttpError as error:
            if json.loads(error.content.decode())['error']['code'] == 429:
                print('Too many requests. Please try again later.')
            else:
                raise
                
      def check_for_spreadsheet_updates(self, spreadsheet_id, last_checked_time):
            service = self.create_service()
            spreadsheet = self.get_spreadsheet(spreadsheet_id)
            current_modified_time = spreadsheet['properties']['modifiedTime']

            if last_checked_time is None or current_modified_time > last_checked_time:
                return True  # Spreadsheet has been changed
            else:
                return False  # No changes since last check

# Example usage
client = GoogleSheetsClient('credentials.json')
spreadsheet_id = 'your_spreadsheet_id'
sheet_name = 'Sheet1'

# Write a value to the next empty cell in a column
next_row = client.get_next_empty_cell_in_column(spreadsheet_id, sheet_name, 'A')
client.write_cell(spreadsheet_id, sheet_name, next_row, 'A', 'New Value')

# Write a value to the next empty cell in a row
next_col = client.get_next_empty_cell_in_row(spreadsheet_id, sheet_name, '1')
client.write_cell(spreadsheet_id, sheet_name, '1', next_col, 'New Value')

# Create a graph in a spreadsheet
graph_data_range = {
    'sheet_name': sheet_name,
    'start_row_index': 1,
    'end_row_index': 10,
    'start_col_index': 1,
    'end_col_index': 2
}
client.create_graph(spreadsheet_id, sheet_name, 'My Graph', graph_data_range, 'LINE')

# Update a graph in a spreadsheet
chart_id = 123456  # ID of the chart to update
client.update_graph(spreadsheet_id, chart_id, graph_data_range)

# Create a new page in a spreadsheet
new_sheet_title = 'New Sheet'
client.create_sheet(spreadsheet_id, new_sheet_title)

# Watch for changes in a specific column or row
callback_url = 'https://your-callback-url.com'
client.watch_changes(spreadsheet_id, callback_url, sheet_name, 'A')
