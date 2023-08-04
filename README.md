# Google Sheets API Python Application

This Python application uses the Google Sheets API to perform various operations on Google Sheets.

## Features

1. **Authentication**: The application authenticates with Google's servers using OAuth 2.0.

2. **Sheet Operations**: The application can perform several operations related to Google Sheets:

   - **Create a new sheet**: The application can create a new sheet in a specific Google Sheets document.
   
   - **Read data**: The application can read data from a specified range in a Google Sheets document.
   
   - **Write data**: The application can write data to a specified range in a Google Sheets document.
   
   - **Delete a sheet**: The application can delete a specific sheet from a Google Sheets document.

   - **Update sheet properties**: The application can update properties of a specific sheet, such as the title and grid properties like row count and column count.

3. **Cell Operations**: The application can perform several operations related to cells in a Google Sheet:

   - **Format cells**: The application can apply formatting to a range of cells in a specific sheet.

   - **Create conditional formatting**: The application can create a conditional formatting rule for a range of cells in a specific sheet.

   - **Set protection**: The application can add a protected range to a specific sheet which restricts users from editing the cells in the range.

   - **Create a filter**: The application can set a basic filter for a range of cells in a specific sheet.

   - **Execute a formula**: The application can execute a formula in a specific cell of a specific sheet.

## Usage

The entry point of the application is the `main.py` file. This file imports functions from the `auth.py`, `sheet_operations.py`, and `cell_operations.py` files to perform tasks.

The user must provide their Google Sheets document ID and the necessary OAuth 2.0 credentials for the application to work. You can find out how to set this up from the official Google documentation: 

https://developers.google.com/workspace/guides/get-started 

## Setup

This application requires Python 3.6 or higher. All required libraries can be installed using next command:

```shell
pip install -r requirements.txt
```

