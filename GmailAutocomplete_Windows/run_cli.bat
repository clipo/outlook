@echo off
echo Gmail to Outlook Autocomplete Builder (CLI)
echo ==========================================
echo.
echo This will scan your Gmail sent messages and create
echo a CSV file for importing into Outlook.
echo.
echo You'll need your Gmail address and an App Password.
echo Get an App Password at: https://myaccount.google.com/apppasswords
echo.
set /p EMAIL="Enter your Gmail address: "
set /p PASSWORD="Enter your App Password: "
echo.
GmailAutocompleteCLI.exe %EMAIL% --password %PASSWORD%
echo.
pause
