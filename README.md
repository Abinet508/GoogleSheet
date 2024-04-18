GoogleSheet
# GoogleSheet

This Python script provides a class `GoogleSheet` that allows you to interact with Google Sheets using the `pygsheets` library. The class provides functionality to read, write, update, and delete data from Google Sheets.

## Purpose

The purpose of this script is to provide an easy-to-use interface for interacting with Google Sheets from Python. This can be useful in a variety of scenarios, such as automating data entry tasks, performing data analysis, or integrating with other Python applications.

## Usage

To use this script, you need to instantiate the `GoogleSheet` class with an `args` object that contains the following properties:

- `credentials`: The path to your Google API credentials file. If the file is not found, the script will prompt you to enter the path manually. If the file still cannot be found, it will default to a file named "wise-logic-415907-d45cefa39f1d.json" in the "credentials" directory.
- `spreadsheet_id`: The ID of the Google Spreadsheet you want to interact with. This is required.
- `worksheet`: The name of the worksheet in the spreadsheet you want to interact with. If not provided, the script will default to the first worksheet.

> Here is an example of how to use the `GoogleSheet` class:

```
args = argparse.Namespace()
args.credentials = 'path/to/credentials.json'
args.spreadsheet_id = 'your_spreadsheet_id'
args.worksheet = 'your_worksheet_name'
gs = GoogleSheet(args)
```

> Please note that this script requires the pygsheets, pandas, os, click, and argparse libraries. You can install these with pip:

`pip install pygsheets pandas click argparse`

