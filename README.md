# Scoutmasters-Council-Camporee-Email-Program
This code is what we use to automate the acceptance/delayed acceptance emails for our camporee

This code will get a sheet which it is connected to (must be a Google Sheet). It will then reach into the sheet, grab all of the headers (to include email recipients), format and email to send to them, and send it.

## How to add this code to your sheet?
To get this code to work, it takes a bit of work.
1. Log into the account which you wish to send the emails from.
2. Create a Google Sheet and transfer all of the data from your Excel file to the Google Sheet.
3. Ensure that the email templates in the script match the headers in the Google Sheet.
4. Update the code to match what is in the Google Sheet.
5. Verify that the emails are sent correctly (preferably a day or so before actually delpoying this code).
6. Save the code in the *AppsScript* editor.
7. Hit run.
