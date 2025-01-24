// Configuration constants
const FROM_EMAIL = 'constdoc@gmail.com'; // Requires "Send As" permission
const DEFAULT_SUBJECT = 'Your Subject Here'; // Customize as needed
const SPREADSHEET_ID = 'your-spreadsheet-id'; // Replace with your actual spreadsheet ID
const EMAILS_TO_SEND_SHEETNAME = 'distributions_to_send'; // Name of the sheet with emails to send
const SENT_HISTORY_SHEETNAME = 'sent_history'; // Name of the sheet to move processed emails


/**
 * Main function that processes the spreadsheet and sends emails based on the data.
 * Reads from 'distributions_to_send' sheet and moves processed rows to 'sent_history'.
 */
function sendEmails() {
    var spreadsheetId = SPREADSHEET_ID; // Replace with your actual spreadsheet ID
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sourceSheet = spreadsheet.getSheetByName(EMAILS_TO_SEND_SHEETNAME);
    var dataRange = sourceSheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var headerIndex = {};
    for (var h = 0; h < headers.length; h++) {
      headerIndex[headers[h]] = h;
    }
  
    // Get or create the 'sent_history' sheet
    var historySheet = spreadsheet.getSheetByName(SENT_HISTORY_SHEETNAME);
    if (!historySheet) {
      historySheet = spreadsheet.insertSheet(SENT_HISTORY_SHEETNAME);
      // Optionally, set up the headers in 'sent_history' sheet
      var historyHeaders = headers.slice(); // Copy headers
      historyHeaders.push('datetime'); // Add 'datetime' column
      historySheet.appendRow(historyHeaders);
    }
  
    // Iterate over the rows in reverse order to handle deletions correctly
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var distributionEmail = row[headerIndex['distribution_email']];
      var additionalEmails = row[headerIndex['additional_emails']];
      var revuSessionInvite = row[headerIndex['revu_session_invite']];
      var templateValues = row[headerIndex['template_values']];
      var emailTemplateUrl = row[headerIndex['email_template']];
      var files = row[headerIndex['files']];
      var subject = row[headerIndex['email_subject']] || DEFAULT_SUBJECT;
  
      try {
        // Validate required data
        if (!distributionEmail && !additionalEmails) {
          Logger.log('Row ' + (i + 1) + ': No recipient email addresses provided.');
          continue; // Skip this row
        }
        if (!emailTemplateUrl) {
          Logger.log('Row ' + (i + 1) + ': No email template URL provided.');
          continue; // Skip this row
        }
  
        // Process the email
        var emailSent = processEmail(
          distributionEmail,
          additionalEmails,
          revuSessionInvite,
          templateValues,
          emailTemplateUrl,
          files,
          subject
        );
  
        // After successful email send, move the row to 'sent_history'
        if (emailSent) {
          moveRowToSentHistory(sourceSheet, historySheet, i + 1, row);
        }
  
      } catch (error) {
        Logger.log('Error processing row ' + (i + 1) + ': ' + error.message);
        console.error('Error processing row ' + (i + 1), error);
      }
    }
  }

/**
 * Processes a single email sending operation.
 * @param {string} distributionEmail - Primary recipient email address(es)
 * @param {string} additionalEmails - Additional recipient email address(es)
 * @param {string} revuSessionInvite - Revu session invite text containing session ID
 * @param {string} templateValues - JSON string of template values
 * @param {string} emailTemplateUrl - Google Drive URL of the email template
 * @param {string} files - Comma-separated list of file URLs to attach
 * @param {string} subject - Email subject line
 * @returns {boolean} - True if email was sent successfully, false otherwise
 */
function processEmail(distributionEmail, additionalEmails, revuSessionInvite, templateValues, emailTemplateUrl, files, subject) {
    var recipientEmails = compileEmailAddresses(distributionEmail, additionalEmails);
    if (!recipientEmails) {
        Logger.log('No valid recipient emails found.');
        return false;
    }

    var emailBody = getEmailBody(emailTemplateUrl, templateValues, revuSessionInvite);
    if (!emailBody) {
        Logger.log('Failed to generate email body.');
        return false;
    }

    var attachments = getAttachments(files);
    return sendEmail(recipientEmails, emailBody, attachments, subject);
    }

/**
 * Combines and validates email addresses from distribution and additional emails.
 * @param {string} distributionEmail - Primary recipient email address(es)
 * @param {string} additionalEmails - Additional recipient email address(es)
 * @returns {string|null} - Comma-separated list of unique email addresses or null if none valid
 */
function compileEmailAddresses(distributionEmail, additionalEmails) {
  var emails = [];

  if (distributionEmail) {
    emails = emails.concat(getEmailAddressesFromString(distributionEmail));
  }

  if (additionalEmails) {
    emails = emails.concat(getEmailAddressesFromString(additionalEmails));
  }

  // Remove duplicates and empty strings
  emails = emails.filter(function(email, index, self) {
    return email && self.indexOf(email) === index;
  });

  return emails.length > 0 ? emails.join(',') : null;
}

/**
 * Extracts email addresses from a string, including handling 'at' and 'dot' notation.
 * @param {string} s - String containing email addresses
 * @returns {string[]} - Array of extracted email addresses
 */
function getEmailAddressesFromString(s) {
  // Remove lines that start with '//' to avoid matching URLs
  s = s.replace(/^\/\/.*/gm, '');

  // Define the regular expression for matching email addresses
  var regex = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\s+at\s+)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\s+dot\s+))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;

  // Find all matches in the string
  var matches = [];
  var match;
  while ((match = regex.exec(s.toLowerCase())) !== null) {
    if (!match[0].startsWith('//')) {
      // Replace ' at ' with '@' and ' dot ' with '.'
      var email = match[0].replace(/\s+(at|dot)\s+/g, function(m, p1) {
        return p1 === 'at' ? '@' : '.';
      });
      matches.push(email);
    }
  }

  return matches;
}

/**
 * Extracts the session ID from the given text using regex patterns.
 * @param {string} text - Text containing the session ID
 * @returns {string|null} - Extracted session ID or null if not found
 */
function parseSessionId(text) {
  if (!text) {
    Logger.log("No revu_session_invite text provided.");
    return null;
  }

  // Define the regular expressions
  var greedyRegex = /\D(\d{3}-\d{3}-\d{3})(?=$|\D)/g;
  var pickyRegex = /\d{3}-\d{3}-\d{3}/;

  var idsInText = [];
  var match;

  // Find all matches using the greedy regex
  while ((match = greedyRegex.exec(text)) !== null) {
    // Extract the session ID using the picky regex
    var idMatch = match[1].match(pickyRegex);
    if (idMatch) {
      idsInText.push(idMatch[0]);
    }
  }

  // Handle no IDs found
  if (idsInText.length === 0) {
    Logger.log("WARNING: No Bluebeam Session ID found. Enter valid session ID text.");
  }

  // Check if all IDs are equal
  var uniqueIds = Array.from(new Set(idsInText));
  if (uniqueIds.length > 1) {
    Logger.log("WARNING: " + uniqueIds.length + " Bluebeam Session ID matches found and some are not the same.\nUsing the first match.");
  }

  // Return the first session ID found or null
  return idsInText.length > 0 ? idsInText[0] : null;
}

/**
 * Generates the email body by applying template values to the HTML template.
 * @param {string} emailTemplateUrl - Google Drive URL of the email template
 * @param {string} templateValues - JSON string of template values
 * @param {string} revuSessionInvite - Revu session invite text containing session ID
 * @returns {string|null} - Generated HTML email body or null if error occurs
 */
function getEmailBody(emailTemplateUrl, templateValues, revuSessionInvite) {
  try {
    var fileId = getFileIdFromUrl(emailTemplateUrl);
    var templateFile = DriveApp.getFileById(fileId);
    var templateContent = templateFile.getBlob().getDataAsString();
    var values = {};
    if (templateValues) {
      values = JSON.parse(templateValues);
    }

    var sessionId = parseSessionId(revuSessionInvite);
    if (sessionId) {
      values['sessionId'] = sessionId;
    } else {
      Logger.log("No session ID found to include in email body.");
    }

    var htmlTemplate = HtmlService.createTemplate(templateContent);
    for (var key in values) {
      if (values.hasOwnProperty(key)) {
        htmlTemplate[key] = values[key];
      }
    }
    var emailBody = htmlTemplate.evaluate().getContent();
    return emailBody;
  } catch (error) {
    Logger.log('Error generating email body: ' + error.message);
    return null;
  }
}
/*
  * Extracts the file ID from a Google Drive URL.
  * @param {string} url The Google Drive URL.
  * @return {string} The file ID.
  * @throws {Error} If the URL is invalid.
*/
/**
 * Extracts file ID from a Google Drive URL.
 * @param {string} url - Google Drive URL
 * @returns {string} - Extracted file ID
 * @throws {Error} If the URL is invalid or file ID cannot be extracted
 */
function getFileIdFromUrl(url) {
  var idMatch = url.match(/[-\w]{25,}/);
  if (idMatch && idMatch[0]) {
    return idMatch[0];
  } else {
    throw new Error('Invalid Google Drive URL: ' + url);
  }
}

/**
 * Retrieves file attachments from Google Drive URLs.
 * @param {string} filesString - Comma-separated list of Google Drive URLs
 * @returns {Blob[]} - Array of file blobs to attach to the email
 */
function getAttachments(filesString) {
  var attachments = [];
  if (filesString) {
    var fileUrls = filesString.split(/[,;]+/).map(function(url) {
      return url.trim();
    });
    fileUrls.forEach(function(fileUrl) {
      try {
        var fileId = getFileIdFromUrl(fileUrl);
        var file = DriveApp.getFileById(fileId);
        attachments.push(file.getBlob());
      } catch (error) {
        Logger.log('Error attaching file from URL ' + fileUrl + ': ' + error.message);
      }
    });
  }
  return attachments;
}

/**
 * Sends an email using Gmail service.
 * @param {string} recipientEmails - Comma-separated list of recipient email addresses
 * @param {string} emailBody - HTML content of the email
 * @param {Blob[]} attachments - Array of file attachments
 * @param {string} subject - Email subject line
 * @returns {boolean} - True if email was sent successfully, false otherwise
 */
function sendEmail(recipientEmails, emailBody, attachments, subject) {
    try {
      var options = {
        htmlBody: emailBody,
        attachments: attachments,
        from: FROM_EMAIL // Requires "Send As" permission
      };
      GmailApp.sendEmail(recipientEmails, subject, '', options);
      Logger.log('Email sent to: ' + recipientEmails);
      return true; // Indicate that the email was sent successfully
    } catch (error) {
      Logger.log('Error sending email to ' + recipientEmails + ': ' + error.message);
      return false;
    }
  }

/**
 * Moves a processed row from source sheet to history sheet and adds timestamp.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - Source spreadsheet sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} historySheet - History spreadsheet sheet
 * @param {number} rowIndex - Index of the row to move
 * @param {Array} rowData - Data from the row being moved
 */
function moveRowToSentHistory(sourceSheet, historySheet, rowIndex, rowData) {
    // Prepare the data to append, adding the timestamp
    var timestamp = new Date();
    var rowToAppend = rowData.slice(); // Copy the row data
    rowToAppend.push(timestamp); // Add timestamp
  
    // Append to 'sent_history' sheet
    historySheet.appendRow(rowToAppend);
  
    // Delete the row from the source sheet
    sourceSheet.deleteRow(rowIndex);
  }
