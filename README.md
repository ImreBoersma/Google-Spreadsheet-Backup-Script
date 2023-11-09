# Spreadsheet Backup Script

This script is designed to create a backup of the current year's data from a Google Spreadsheet. The code retrieves the active spreadsheet, identifies the current year's sheet, creates a backup sheet, and copies the data into the newly created backup.

## Prerequisites

- This script is written in JavaScript and designed to run within the Google Apps Script environment.
- Ensure the necessary permissions are granted within your Google account to access and manipulate spreadsheets.

## Instructions

1. Copy the code provided and paste it into your Google Apps Script editor.
2. Save the script and run the `main()` function to execute the backup process.

## Disclaimer

This script assumes the existence of a sheet named 'Huidig jaar' (Current year) in the active spreadsheet. If the sheet doesn't exist, the script will throw an error.

Always test scripts in a safe environment to ensure they function as intended before applying them to important data.
