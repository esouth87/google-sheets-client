# google-sheets-client
 Google Sheets Python Library: A Python library for interacting with Google Sheets. Edit cells, create/update graphs, add new sheets, and monitor changes. Simplify your Google Sheets automation.
# Google Sheets Python Library

The Google Sheets Python Library provides a simple and convenient way to interact with Google Sheets using the Google Sheets API. It allows you to edit cells, create and update graphs, create new sheets, and monitor changes in specified columns or rows.

## Features

- Edit the next empty cell in a column
- Edit the next empty cell in a row
- Create graphs in a spreadsheet
- Update and edit graphs in a spreadsheet
- Create a new page in a spreadsheet
- Watchdog function to trigger an action when a particular column or row is changed

## Getting Started

1. Set up a Google Cloud Platform project and obtain the necessary credentials.
2. Install the required dependencies using pip: `pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client`.
3. Import the `GoogleSheetsClient` class from the library and instantiate it with your credentials.
4. Use the provided methods to interact with Google Sheets.

## Example Usage

```python
# Instantiate the client
client = GoogleSheetsClient('credentials.json')

# Write a value to the next empty cell in a column
client.write_cell(spreadsheet_id, sheet_name, row, col, value)

# Write a value to the next empty cell in a row
client.write_cell(spreadsheet_id, sheet_name, row, col, value)

# Create a graph in a spreadsheet
client.create_graph(spreadsheet_id, sheet_name, graph_title, data_range, graph_type)

# Update a graph in a spreadsheet
client.update_graph(spreadsheet_id, chart_id, data_range)

# Create a new page in a spreadsheet
client.create_sheet(spreadsheet_id, sheet_title)

# Watch for changes in a specific column or row
client.watch_changes(spreadsheet_id, callback_url, sheet_name, row_or_col)
