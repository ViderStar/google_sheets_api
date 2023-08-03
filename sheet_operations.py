from main import SAMPLE_SPREADSHEET_ID

def create_sheet(service, title):
    """Create a new sheet with the provided title and return its properties."""
    try:
        sheet = service.spreadsheets()
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': title
                    }
                }
            }]
        }
        result = sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
        return result['replies'][0]['addSheet']['properties']
    except Exception as e:
        print(f"Failed to create sheet: {e}")
        return None


def read_data(service, range_name):
    """Read and return data from the specified range of the sheet."""
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=range_name).execute()
        values = result.get('values', [])
        return values
    except Exception as e:
        print(f"Failed to read data from sheet: {e}")
        return None


def write_data(service, range_name, values):
    """Write the provided data to the specified range of the sheet."""
    try:
        sheet = service.spreadsheets()
        body = {'values': values}
        result = sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
        return result.get("updatedCells")
    except Exception as e:
        print(f"Failed to write data to sheet: {e}")
        return None


def update_sheet_properties(service, sheet_id, title, grid_properties):
    """
    Updates the properties of a specific sheet. The properties that can be updated include the title of the
    sheet and the grid properties like row count and column count.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'title': title,
                        'gridProperties': grid_properties
                    },
                    'fields': 'title,gridProperties'
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result


def delete_sheet(service, sheet_id):
    """Delete the sheet with the provided sheet id."""
    try:
        sheet = service.spreadsheets()
        body = {
            'requests': [{
                'deleteSheet': {
                    'sheetId': sheet_id
                }
            }]
        }
        sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
        return sheet_id
    except Exception as e:
        print(f"Failed to delete sheet: {e}")
        return None
