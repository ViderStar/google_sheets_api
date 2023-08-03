def format_cells(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index):
    """
    Applies formatting to a range of cells in a specific sheet. The formatting includes background color,
    text alignment, text color, font size and boldness.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': start_row_index,
                        'endRowIndex': end_row_index,
                        'startColumnIndex': start_column_index,
                        'endColumnIndex': end_column_index,
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red': 1.0,
                                'green': 0.0,
                                'blue': 0.0
                            },
                            'horizontalAlignment': 'CENTER',
                            'textFormat': {
                                'foregroundColor': {
                                    'red': 1.0,
                                    'green': 1.0,
                                    'blue': 1.0
                                },
                                'fontSize': 12,
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result


def create_conditional_formatting(service, sheet_id, start_row_index, end_row_index, start_column_index,
                                  end_column_index):
    """
    Creates a conditional formatting rule for a range of cells in a specific sheet. The rule changes the
    background color of the cell if the value in the cell is greater than 10.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [
                            {
                                'sheetId': sheet_id,
                                'startRowIndex': start_row_index,
                                'endRowIndex': end_row_index,
                                'startColumnIndex': start_column_index,
                                'endColumnIndex': end_column_index
                            }
                        ],
                        'booleanRule': {
                            'condition': {
                                'type': 'NUMBER_GREATER',
                                'values': [
                                    {
                                        'userEnteredValue': '10'
                                    }
                                ]
                            },
                            'format': {
                                'backgroundColor': {
                                    'red': 0.0,
                                    'green': 1.0,
                                    'blue': 0.0
                                }
                            }
                        }
                    },
                    'index': 0
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result


def set_protection(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index):
    """
    Adds a protected range to a specific sheet. The protected range restricts the users from editing the cells
    in the range. The protection can be set to warning only mode allowing users to edit after confirming.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'addProtectedRange': {
                    'protectedRange': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': start_row_index,
                            'endRowIndex': end_row_index,
                            'startColumnIndex': start_column_index,
                            'endColumnIndex': end_column_index
                        },
                        'description': 'Protected Range',
                        'warningOnly': True  # Change this to False for stricter protection
                    }
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result


def create_filter(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index):
    """
    Sets a basic filter for a range of cells in a specific sheet. The filter allows users to filter rows in
    the range based on the cell values.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'setBasicFilter': {
                    'filter': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': start_row_index,
                            'endRowIndex': end_row_index,
                            'startColumnIndex': start_column_index,
                            'endColumnIndex': end_column_index
                        }
                    }
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result

from main import SAMPLE_SPREADSHEET_ID

def execute_formula(service, sheet_id, row, column, formula):
    """
    Executes a formula in a specific cell of a specific sheet. The formula to be executed is provided as input.
    """
    sheet = service.spreadsheets()
    body = {
        'requests': [
            {
                'updateCells': {
                    'rows': [
                        {
                            'values': [
                                {
                                    'userEnteredValue': {
                                        'formulaValue': formula
                                    }
                                }
                            ]
                        }
                    ],
                    'fields': 'userEnteredValue.formulaValue',
                    'start': {
                        'sheetId': sheet_id,
                        'rowIndex': row,
                        'columnIndex': column
                    }
                }
            }
        ]
    }
    result = sheet.batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
    return result
