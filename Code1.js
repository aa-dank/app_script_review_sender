
function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/)[0];
}

function sendEmail(recipients, htmlBody, attachments) {
  GmailApp.sendEmail(recipients.join(','), 'Subject', '', {
    htmlBody: htmlBody,
    attachments: attachments,
    from: 'constdoc@gmail.com'
  });
}

function parseEmails(distributionEmail, additionalEmails) {
  const allEmails = distributionEmail + ',' + additionalEmails;
  return getEmailAddresses(allEmails);
}

function getEmailAddresses(s) {
  // Regular expression for matching email addresses
  const regex = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\sat\s)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\sdot\s))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;

  // Convert input to lowercase
  s = s.toLowerCase();

  // Find all matches
  const matches = s.match(regex) || [];

  // Filter out matches starting with '//' and return the result
  return matches.filter(email => !email.startsWith('//'));
}

function getAttachments(filesString) {
  const fileUrls = filesString.split(',');
  return fileUrls.map(url => DriveApp.getFileById(getIdFromUrl(url)).getBlob());
}

function getFormattedEmailBody(templateUrl, templateValues) {
  const template = DriveApp.getFileById(getIdFromUrl(templateUrl)).getBlob().getDataAsString();
  const parsedValues = JSON.parse(templateValues);
  
  // Use Mustache.js or similar templating library
  return Mustache.render(template, parsedValues);
}

function processRow(row) {
  const [distributionEmail, additionalEmails, revuSessionInvite, templateValues, emailTemplate, files] = row;
  
  const recipients = parseEmails(distributionEmail, additionalEmails);
  const htmlBody = getFormattedEmailBody(emailTemplate, templateValues);
  const attachments = getAttachments(files);
  
  sendEmail(recipients, htmlBody, attachments);
}

function sendDistributionEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('distributions_to_send');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Skip header row
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    processRow(row);
  }
}

function myFunction() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("distributions_to_send")
    
}

sendDistributionEmails();
