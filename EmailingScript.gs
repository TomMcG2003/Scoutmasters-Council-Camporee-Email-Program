/*
 * Scoutmasters' Council Email Automation
 * Copyright (C) 2025 onward, Thomas M. McGowan <tmcgowan2025@gmail.com>
 */

function sendEmailsFromSheet(sheetName, emailSubject, emailTemplate) {
  /**
   * const sheetName: This is the name of the sheet for which you would like to access the data for
   * const emailSubject: This is the subject of the email you wish to send
   * const emailTemplate: This is the shell of the email you would like to send. Ensure that the script is
   * correctly formatted to accept all of the requisite data.
   * 
   * Output: This function will send the emails to the address listed in the sheet.
   *  However, it will not handle invlaid emails (it will still register them as sent)
   *  and it can only send < 500 emails per day
   */
  // Access the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  // Handles the case that the sheet is not there.
  if (!sheet) {
    Logger.log(`Sheet with name '${sheetName}' not found.`);
    return;
  }
  
  // Get the data from the specified sheet
  const data = sheet.getDataRange().getValues();
  // First row contains headers
  const headers = data[0]; 
  // Exclude the header row
  const rows = data.slice(1); 
  
  // Send emails
  rows.forEach(row => {
    const emailData = {};  // Initializes the email data
    // This will collect the data for that row.
    headers.forEach((header, index) => {
      emailData[header] = row[index]; // Map each column to its corresponding header
    });
    // Combine the email addresses
    const recipients = [emailData["POC Email"], emailData["Alternate POC Email"]]
      .filter(email => email) // Remove any empty email fields
      .join(","); // Separate multiple recipients with commas
    
    // Replace placeholders in the email template
    const personalizedEmail = emailTemplate.replace(/{{(.*?)}}/g, (match, p1) => emailData[p1.trim()] || "");
    
    // Send email if there are recipients
    if (recipients) {
      // This try/catch statement will log erros and send them to the console.
      try{
        GmailApp.sendEmail(recipients, emailSubject, personalizedEmail);
      } catch(e){
        Logger.log(`Failed to send the email to ${recipients}, the error is ${e}`);
      }

      Logger.log("Sending email");
    }
    else{
      Logger.log("No recipients");
    }
  }
  
  );
}

function sendEmailsFromSheetWithFormatValidation(sheetName, emailSubject, emailTemplate){
  // Access the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  // Handles the case that the sheet is not there.
  if (!sheet) {
    Logger.log(`Sheet with name '${sheetName}' not found.`);
    return;
  }
  // Gets the data from the Google sheet
  const data = sheet.getDataRange().getValues();
  // Grabs the headers of the columns
  const headers = data[0];
  // Grabs the rest of the data and splits it into rows
  const rows = data.slice(1);
  // This is the regular expression that we will be working with to validate email address formats
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Basic email validation regex
  // TODO: Update this regex if you would like more advanced email validation.

  /**
   * This is a for loop. What it does is go through every row in the Google Sheet 
   * and parses the data by the headers. It will then store the email recipients
   * in the <recipients> variable. This serves as our "to" line. It will then 
   * validate the emails using the regular expression and log any non-valid emails
   * in the console that pops up when you hit "run". It will then parse the email 
   * script and add the personalized data to it. Given that there is at least one
   * valid recipient, it will send the email to them. All external actions will be 
   * logged in the console. 
   */
  rows.forEach(row =>{
    const emailData = {};
    headers.forEach((header, index) => {
      emailData[header] = row[index];
    });
    // This line will go into the data of the Google Sheet and grab the email addresses.
    const recipients = [emailData["POC Email"], emailData["Alternate POC Email"]]
    // TODO: Update the strings in the quotations to reflect the correct headers in the Google Sheet
    .filter(email => email && emailRegex.test(email))
    .join(",");
    // This if statement captures any instance where there are no valid email addresses
    if (!recipients){
      Logger.log(`No valid email addresses for row: ${JSON.stringify(emailData)}`);
      return;
    }
    // This line is where we input the data from the Google Sheet to personalize it
    const personalizedEmail = emailTemplate.replace(/{{(.*?)}}/g, (match, p1) => emailData[p1.trim()] || "");
    // This try/catch will handle any errors that we get and log them to the console.
    try{
      GmailApp.sendEmail(recipients, emailSubject, personalizedEmail);
      Logger.log(`Email number ${index} sent to ${recipients}; Troop ID: ${emailData["Troop ID"]}`);
    } catch (e){
      Logger.log(`Error sending email number ${index} to ${recipients}, Troop ID: ${emailData["Troop ID"]}. Error is ${e}`);
    }

  });
  
}

const emailSubject = "West Point Scoutmasters' Council 61st Camporee";

// TODO: Change the variables in the "{{}}" to match what is in the Google Sheet
const acceptedShell = `
    Troop {{Troop Number}}, 

On behalf of West Point’s Scoutmasters’ Council, I am pleased to inform you that you have been selected to attend this year’s camporee on April 11th – 13th. Congratulations! 

Below is further information on your selection and allotted slots:  

Unique Camporee ID:  {{Unique ID}}

Troop Number: {{Troop Number}}

Council: {{Council}}

City: {{City}}

State: {{State}}

Camporee Group: {{Camporee Group}} 

Attendees: {{Attendees}} 

 

The first information packet will be sent to you soon. Many of your questions will hopefully be answered in this packet. However, if you have any urgent questions, please feel free to reach out. 

We look forward to welcoming you to this year’s camporee. 

Sincerely,
West Point's Scoutmasters' Council
  `;
// TODO: Change the variables in the "{{}}" to match what is in the Google Sheet
const delayedShell = `
Troop {{Troop Number}}, 

On behalf of West Point’s Scoutmasters’ Council, we would like to thank you for your application. Due to constraints, we are not allowed to accept every applicant. We would like to offer you the opportunity to attend next year’s camporee on a delayed acceptance.  

Delayed Acceptance ID:  {{Unique ID}}

Total Number Allowed:
  To avoid issues with acceptances next year, we have to limit the number of Scouts that you can apply for next year. You are only allowed to apply for a MAXIMUM of {{Attendees}} next year. You are able to bring less than {{Attendees}}, though no more than that. Thank you for your understanding.

You will be able to input this delayed acceptance ID into the 2026 Camporee Registration form to be selected for the Camporee for 2026. 
`;

function sendAllEmails(){
sendEmailsFromSheet("Acceptances", emailSubject, acceptedShell);
sendEmailsFromSheet("Delayed", emailSubject, delayedShell);
}

function sendAllEmailsWithValidation(){
  sendAllEmailsWithValidation("Acceptances", emailSubject, acceptedShell);
  sendAllEmailsWithValidation("Delayed", emailSubject, delayedShell);
}

// sendAllEmails();
// sendAllEmailsWithValidation();

Logger.log(`Sent all emails.`);

