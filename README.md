# Scoutmasters-Council-Camporee-Email-Program
This code is what we use to automate the acceptance/delayed acceptance emails for our camporee

This code will get a sheet which it is connected to (must be a Google Sheet). It will then reach into the sheet, grab all of the headers (to include email recipients), format and email to send to them, and send it.

## How to add this code to your sheet?
To get this code to work, it takes a bit of work.
1. Log into the account which you wish to send the emails from.
2. Create a Google Sheet and transfer all of the data from your Excel file to the Google Sheet.
3. On the Google Sheet, click the "Extensions" button in the navigation bar and then select the "Apps Script" button. You should now be in a code editor (I am sorry that there is no dark mode for this).
4. Ensure that the email templates in the script match the headers in the Google Sheet.
5. Update the code to match what is in the Google Sheet.
6. Verify that the emails are sent correctly (preferably a day or so before actually delpoying this code).
7. Save the code in the *AppsScript* editor.
8. Hit run.

## Notes:
- Read all the TODOs in the code. This will guide you on what you need to change for the code to work with your specific data.
- Read all the comments throughout. This will help you get an understanding of what the code does. This also allows you to find any errors that you may run into and fix them!

  ## Point of Contact:
  The point of contact for this code is Thomas McGowan. He can be reached by email at tmcgowan2025@gmail.com. Please report any bugs to him or open an issue here on GitHub.
