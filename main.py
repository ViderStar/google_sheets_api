from auth import create_service
from sheet_operations import create_sheet, read_data, write_data, delete_sheet, update_sheet_properties
from cell_operations import format_cells, create_conditional_formatting, set_protection, create_filter, execute_formula

SAMPLE_SPREADSHEET_ID = '1B81-Kfgnqs8KEZKTCsW410nvUsFbWIDlbBafwIdpA8s'

if __name__ == '__main__':
    service = create_service()

    # Read data
    SAMPLE_RANGE_NAME = 'Sheet1!A1:E5'
    values = read_data(service, SAMPLE_RANGE_NAME)
    if values:
        print('Read values:')
        for row in values:
            print(row)

    # Write data
    values = [['New data 1', 'New data 2']]
    updated_cells = write_data(service, SAMPLE_RANGE_NAME, values)
    print(f'Written {updated_cells} cells')

    # Create a new sheet
    new_sheet_properties = create_sheet(service, 'New Sheet')
    print('Created new sheet:', new_sheet_properties['title'])

    # Format cells
    sheet_id = new_sheet_properties['sheetId'] # Using the newly created sheet's ID
    start_row_index = 0
    end_row_index = 5
    start_column_index = 0
    end_column_index = 5
    format_cells(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index)

    # Create conditional formatting
    create_conditional_formatting(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index)

    # Set protection
    set_protection(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index)

    # Create filter
    create_filter(service, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index)

    # Execute formula
    row = 0
    column = 0
    formula = '=A1+B1'
    execute_formula(service, sheet_id, row, column, formula)

    # Update sheet properties
    title = 'New Sheet Title'
    grid_properties = {'rowCount': 100, 'columnCount': 20}
    update_sheet_properties(service, sheet_id, title, grid_properties)

    # Delete a sheet
    # deleted_sheet_id = delete_sheet(service, new_sheet_properties['sheetId'])
    # print('Deleted sheet with ID:', deleted_sheet_id)