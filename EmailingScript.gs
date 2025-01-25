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
    // TODO: You need to change the strings in the quotations so that they match what is in the Google Sheet.
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

const emailSubject = "West Point Scoutmasters' Council 61st Camporee";
// TODO: Update this subject so that it is the current year's camporee

// TODO: Update all of the strings that are within the '{{}}' so that they match the headers in the Google Sheet
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
// TODO: Update all of the strings that are within the '{{}}' so that they match the headers in the Google Sheet
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
sendAllEmails();
