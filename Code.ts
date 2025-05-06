/**
 * @fileoverview Email distribution system for review document management.
 * This script handles the automated sending of email distributions with attachments
 * using Google Apps Script. It processes data from a spreadsheet, sends emails with
 * customizable templates, and tracks sent distributions.
 * 
 * @author 
 * @version 1.0.0
 */

/** Defines the sheet names used in the spreadsheet */
enum SheetNames {
  /** Sheet containing pending email distributions */
  TO_SEND = 'distributions_to_send',
  /** Sheet containing record of sent distributions */
  SENT_HISTORY = 'sent_history',
  /** Sheet with sets of template values used to populate TO_SEND */
  TEMPLATES = 'distribution_templates'
}

/**
 * Global configuration interface for the email distribution system
 * @interface Config
 */
interface Config {
  /** Email address to send from (requires Gmail "Send As" permissions) */
  readonly FROM_EMAIL: string;
  /** Default subject line for emails when none is provided */
  readonly DEFAULT_SUBJECT: string;
  /** ID of the Google Spreadsheet containing distribution data */
  readonly SPREADSHEET_ID: string;
  /** Enable debug logging */
  readonly DEBUG_MODE: boolean;
}

const CONFIG: Config = {
  FROM_EMAIL: 'constdoc@ucsc.edu',
  DEFAULT_SUBJECT: 'Your Subject Here',
  // prod spreadsheet id: 1bcbDMbwdch6BE1KzBrUPi0f4FQHiVd-ya8kPcrnnQ5Q
  SPREADSHEET_ID: '1RfbiEpwU2APw3fXg5VoD4Dg2PlvacTV7EIrb_h98mcY',
  DEBUG_MODE: true
};

const MAX_ATTACHMENT_SIZE = 21 * 1024 * 1024; // 21MB

/**
 * Enhanced logging utility
 */
class CustomLogger {
  /**
   * Logs a debug message if debug mode is enabled
   * @param {string} message - Debug message
   * @param {any} [data] - Optional data to include in the log
   */
  static debug(message: string, data?: any) {
    if (CONFIG.DEBUG_MODE) {
      Logger.log(`[DEBUG] ${message} ${data ? JSON.stringify(data) : ''}`);
    }
  }

  /**
   * Logs an error message with detailed error information
   * @param {string} message - Error description
   * @param {any} error - Error object or details
   */
  static error(message: string, error: any) {
    Logger.log(`[ERROR] ${message}`);
    
    // Handle different error formats
    const errorDetails = error instanceof Error 
      ? {
          message: error.message,
          name: error.name,
          stack: error.stack,
          toString: error.toString()
        }
      : error;
    
    Logger.log('Error details:' + JSON.stringify(errorDetails));
  }

  /**
   * Logs an info message regardless of debug mode
   * @param {string} message - Info message
   */
  static info(message: string) {
    Logger.log(`[INFO] ${message}`);
  }
}

// Types and Interfaces
/**
 * Represents a row of email distribution data from the spreadsheet
 * @interface EmailRow
 */
interface EmailRow {
  /** Primary distribution list email addresses */
  distribution_emails: string;
  /** Additional individual email addresses to include */
  additional_emails: string;
  /** Bluebeam Revu session invite text containing session ID */
  revu_session_invite: string;
  /** JSON string containing template variable values */
  template_values: string; // Updated property name
  /** Google Drive URL of the email template HTML file */
  email_body_template: string;
  /** Comma-separated list of Google Drive URLs for attachments */
  attachments_urls: string; // Changed property name from "files"
  /** Custom subject line for the email */
  email_subject_template: string;
  subject_template_value: string;
  /** Template label to identify which distribution template to use */
  template_label: string;
  /** Allow for additional dynamic columns */
  [key: string]: string; // Allow additional columns
}

interface ProcessedEmailRow extends EmailRow {
  datetime: Date;
}

interface TemplateValues {
  [key: string]: string;
  sessionId?: string;
}

/**
 * Utility class for spreadsheet operations
 */
class SpreadsheetUtils {
  /**
   * Maps headers from a sheet to their column indices
   * @param {any[]} headerRow - First row of a sheet containing header names
   * @returns {Object} Map of header names to their column indices
   */
  static mapHeadersToIndices(headerRow: any[]): { [key: string]: number } {
    return headerRow.reduce((acc: { [key: string]: number }, header: string, index: number) => {
      acc[header] = index;
      return acc;
    }, {});
  }

  /**
   * Gets or creates a sheet in the spreadsheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
   * @param {string} sheetName - Name of the sheet to get or create
   * @param {string[]} [headers] - Optional headers to add if the sheet is created
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
   */
  static getOrCreateSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string, headers?: string[]): GoogleAppsScript.Spreadsheet.Sheet {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      if (headers && headers.length > 0) {
        sheet.appendRow(headers);
        CustomLogger.debug(`Created sheet ${sheetName} with headers`);
      }
    }
    return sheet;
  }

  /**
   * Ensures a column exists in a sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to check
   * @param {string} columnName - The column name to look for
   * @returns {boolean} True if the column was added, false if it already existed
   */
  static ensureColumnExists(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnName: string): boolean {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes(columnName)) {
      const newColIndex = headers.length + 1;
      sheet.getRange(1, newColIndex).setValue(columnName);
      CustomLogger.debug(`Added ${columnName} column to ${sheet.getName()} sheet`);
      return true;
    }
    CustomLogger.debug(`${columnName} column already exists in ${sheet.getName()} sheet`);
    return false;
  }

  /**
   * Maps a data row to an object using header indices
   * @param {any[]} row - The data row
   * @param {{ [key: string]: number }} headers - Map of header names to column indices
   * @param {string[]} requiredFields - Fields that should always exist in the result
   * @returns {Object} Object with header names as keys and row values as values
   */
  static mapRowToObject(row: any[], headers: { [key: string]: number }, requiredFields: string[]): { [key: string]: any } {
    const result: { [key: string]: any } = {};
    
    // Add required fields first to ensure they exist
    for (const field of requiredFields) {
      result[field] = headers[field] !== undefined ? row[headers[field]] || '' : '';
    }
    
    // Add any additional columns from the row data
    for (const header in headers) {
      if (!Object.prototype.hasOwnProperty.call(result, header)) {
        result[header] = row[headers[header]] || '';
      }
    }
    
    return result;
  }
}

/**
 * Utility class for file operations
 */
class FileUtils {
  /**
   * Extracts a file ID from a Google Drive URL
   * @param {string} url - Google Drive URL
   * @returns {string} The extracted file ID
   * @throws {Error} If the URL is invalid or file ID cannot be extracted
   */
  static extractFileId(url: string): string {
    const match = url.match(/[-\w]{25,}/);
    if (!match?.[0]) {
      throw new Error(`Invalid Google Drive URL: ${url}`);
    }
    return match[0];
  }

  /**
   * Gets file content from a Google Drive file URL
   * @param {string} url - Google Drive URL
   * @returns {string} The file content as string
   */
  static getFileContentFromUrl(url: string): string {
    const fileId = this.extractFileId(url);
    CustomLogger.debug(`Retrieving content for fileId: ${fileId}`);
    
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    
    CustomLogger.debug(`Content retrieved (first 100 chars): ${content.substring(0, 100)}`);
    return content;
  }

  /**
   * Checks if a file exceeds the maximum allowed size
   * @param {string} fileId - Google Drive file ID
   * @param {number} maxSize - Maximum allowed size in bytes
   * @returns {boolean} True if file exceeds the size limit
   */
  static isFileTooLarge(fileId: string, maxSize: number): boolean {
    const file = DriveApp.getFileById(fileId);
    return file.getSize() > maxSize;
  }

  /**
   * Gets file as a blob from its Google Drive ID
   * @param {string} fileId - Google Drive file ID
   * @returns {GoogleAppsScript.Base.Blob} The file blob
   */
  static getFileBlob(fileId: string): GoogleAppsScript.Base.Blob {
    return DriveApp.getFileById(fileId).getBlob();
  }

  /**
   * Moves a file to trash
   * @param {string} fileId - Google Drive file ID
   * @returns {boolean} True if successful
   */
  static trashFile(fileId: string): boolean {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      return true;
    } catch (error) {
      CustomLogger.error(`Error trashing file: ${fileId}`, error);
      return false;
    }
  }

  /**
   * Gets file metadata for debugging purposes
   * @param {string} fileId - Google Drive file ID
   * @returns {object} File metadata
   */
  static getFileMetadata(fileId: string): Record<string, any> {
    try {
      const file = DriveApp.getFileById(fileId);
      return {
        id: file.getId(),
        name: file.getName(),
        owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
        sharingAccess: file.getSharingAccess(),
        size: file.getSize()
      };
    } catch (error) {
      return { id: fileId, error: error.toString() };
    }
  }
}

/**
 * Utility class for email operations
 */
class EmailUtils {
  /**
   * Parses email addresses from text, handling both standard format and 'at'/'dot' notation
   * 
   * This method handles various formats of email addresses, including:
   * - Standard email addresses (user@example.com)
   * - Email addresses with 'at' notation (user at example.com)
   * - Email addresses with 'dot' notation (user@example dot com)
   * - Combined formats (user at example dot com)
   * 
   * It also filters out lines that start with '//' to prevent matching URLs or comments.
   * 
   * The regex matches email local parts that can contain:
   * - Alphanumeric characters
   * - Special characters: !#$%&'*+/=?^_`{|}~-
   * - Dots between parts
   * 
   * Followed by @ or " at " (case insensitive), followed by:
   * - Domain name parts (with - allowed within parts)
   * - Dots or " dot " between parts
   * 
   * @param {string} input - Text containing email addresses in various formats
   * @returns {string[]} Array of extracted and normalized email addresses
   * @example
   * // Returns: ["john.doe@example.com", "jane.smith@company.org"]
   * EmailUtils.parseEmailAddresses("Contact: john.doe@example.com or jane.smith at company dot org");
   */
  static parseEmailAddresses(input: string): string[] {
    if (!input) return [];
    
    // Remove comment lines starting with '//' to avoid matching URLs
    const cleanInput = input.replace(/^\/\/.*/gm, '').toLowerCase();
    
    // Complex regex to match various email formats
    const emailRegex = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\s+at\s+)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\s+dot\s+))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;

    const matches: string[] = [];
    let match: RegExpExecArray | null;
    
    while ((match = emailRegex.exec(cleanInput)) !== null) {
      if (!match[0].startsWith('//')) {
        // Replace ' at ' with '@' and ' dot ' with '.' to normalize the email format
        const email = match[0].replace(/\s+(at|dot)\s+/g, (m, p1) => 
          p1 === 'at' ? '@' : '.'
        );
        matches.push(email);
      }
    }

    return matches;
  }

  /**
   * Extracts Bluebeam session ID from text using regex pattern matching
   * 
   * This method searches for and extracts Bluebeam session IDs from text.
   * A valid session ID follows the pattern: three digits, hyphen, three digits, hyphen, three digits.
   * Example: 123-456-789
   * 
   * Notes on the session ID extraction process:
   * - The regex looks for the specific pattern \b\d{3}-\d{3}-\d{3}\b
   * - \b ensures we're matching whole word boundaries (not part of a larger number)
   * - If multiple session IDs are found, the first one is used
   * - If no valid session ID is found, null is returned
   * - Detailed logs are generated to aid debugging of session ID extraction issues
   * 
   * @param {string} text - Text containing a Bluebeam session ID
   * @returns {string | null} Session ID or null if not found
   * @example
   * // Returns: "123-456-789"
   * EmailUtils.parseSessionId("Please join session 123-456-789 for the review");
   * 
   * // Returns: null
   * EmailUtils.parseSessionId("No valid session ID in this text");
   */
  static parseSessionId(text: string): string | null {
    if (!text) return null;

    // Regular expression to match session IDs in the format '123-456-789'
    const sessionIdRegex = /\b\d{3}-\d{3}-\d{3}\b/;
    const matches = text.match(sessionIdRegex);

    if (!matches) {
      CustomLogger.debug(`No Bluebeam Session ID found in the text: ${text}`);
      return null;
    }

    // Filter out any falsy values from matches
    const ids = matches.filter(Boolean);
    CustomLogger.debug(`Found session IDs: ${ids.join(', ')}`);

    if (ids.length > 1) {
      CustomLogger.debug(`Multiple different session IDs found in the text: ${text}\n Using first match.`);
    }

    return ids[0] || null;
  }

  /**
   * Combines email addresses from multiple sources and removes duplicates
   * 
   * This method:
   * 1. Takes multiple input strings that may contain email addresses
   * 2. Parses each input to extract valid email addresses
   * 3. Combines all found email addresses into a single collection
   * 4. Removes duplicates and filters out any empty/invalid entries
   * 5. Joins valid addresses with commas for use in email sending
   * 
   * @param {string[]} sources - Array of strings containing email addresses
   * @returns {string | null} Comma-separated email list or null if none found
   * @example
   * // Returns: "user1@example.com,user2@example.com,user3@example.com"
   * EmailUtils.combineEmailAddresses(
   *   "user1@example.com, user2@example.com",
   *   "user2@example.com, user3@example.com"
   * );
   */
  static combineEmailAddresses(...sources: string[]): string | null {
    const allEmails: string[] = [];
    
    for (const source of sources) {
      if (source) {
        allEmails.push(...this.parseEmailAddresses(source));
      }
    }

    const uniqueEmails = [...new Set(allEmails.filter(Boolean))];
    return uniqueEmails.length > 0 ? uniqueEmails.join(',') : null;
  }

  /**
   * Sends an email using Gmail service
   * @param {string} recipients - Comma-separated recipient email addresses
   * @param {string} subject - Email subject
   * @param {string} htmlBody - HTML content for email body
   * @param {GoogleAppsScript.Base.Blob[]} attachments - Email attachments
   * @param {string} fromEmail - Sender email address
   * @returns {boolean} True if email sent successfully
   */
  static sendEmail(
    recipients: string, 
    subject: string, 
    htmlBody: string, 
    attachments: GoogleAppsScript.Base.Blob[] = [],
    fromEmail: string = CONFIG.FROM_EMAIL
  ): boolean {
    CustomLogger.debug(`Sending email to: ${recipients}, Subject: ${subject}, Attachments: ${attachments.length}`);
    
    try {
      GmailApp.sendEmail(recipients, subject, '', {
        htmlBody,
        attachments,
        from: fromEmail
      });
      CustomLogger.debug('Email sent successfully');
      return true;
    } catch (error) {
      CustomLogger.error('Error sending email', error);
      return false;
    }
  }
}

/**
 * Utility class for text processing
 */
class TextUtils {
  /**
   * Decodes HTML entities to their corresponding characters
   * @param {string} text - Text with HTML entities
   * @returns {string} Decoded text
   */
  static decodeHtmlEntities(text: string): string {
    if (!text) return text;
    return text
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }

  /**
   * Sanitizes and formats JSON text from spreadsheet cells for use in templates
   * 
   * This method handles several common issues with JSON stored in spreadsheet cells:
   * 1. Removes extra escaping that occurs when JSON is pasted into cells
   * 2. Handles newlines that might break JSON parsing
   * 3. Ensures proper JSON structure with curly braces
   * 4. Validates JSON structure through parse/stringify cycle
   * 
   * Common input issues addressed:
   * - Double-escaped quotes (e.g., \\" instead of ")
   * - Missing outer curly braces
   * - Line breaks in JSON text
   * - Inconsistent formatting
   * 
   * @param {string} text - Potentially invalid or malformed JSON text from spreadsheet
   * @returns {string} Valid, properly formatted JSON string or '{}' if input cannot be parsed
   * @example
   * // Returns: {"name":"John","age":"30"}
   * TextUtils.sanitizeJsonText('name: "John", age: "30"');
   * 
   * // Returns: {"project":"Building A","status":"In Progress"}
   * TextUtils.sanitizeJsonText('{"project":"Building A",\n"status":"In Progress"}');
   */
  static sanitizeJsonText(text: string): string {
    if (!text) return text;
    
    try {
      // Remove extra escaping
      let cleaned = text
        .replace(/\\"/g, '"')     // Remove escaped quotes
        .replace(/\\\\/g, '\\')   // Remove double escapes
        .replace(/\r?\n/g, ' ')   // Replace newlines with spaces
        .trim();                  // Remove extra whitespace
        
      // If it's not wrapped in curly braces, wrap it
      if (!cleaned.startsWith('{')) cleaned = '{' + cleaned;
      if (!cleaned.endsWith('}')) cleaned = cleaned + '}';
      
      // Validate by parsing and re-stringifying
      const parsed = JSON.parse(cleaned);
      return JSON.stringify(parsed);
    } catch (error) {
      CustomLogger.error('Error sanitizing JSON text', error);
      return '{}';
    }
  }

  /**
   * Gets a string value safely from an array using a header map
   * @param {any[]} row - Row data array
   * @param {{ [key: string]: number }} headerMap - Map of headers to indices
   * @param {string} headerName - The header name to get value for
   * @returns {string} The value as string or empty string if not found
   */
  static getStringValue(row: any[], headerMap: { [key: string]: number }, headerName: string): string {
    if (headerMap[headerName] === undefined || row[headerMap[headerName]] === undefined) {
      return '';
    }
    return String(row[headerMap[headerName]] || '');
  }
}

/**
 * Handles the construction and sending of individual emails
 * Processes templates, attachments, and email addresses
 */
class EmailBuilder {
  /**
   * Creates a new EmailBuilder instance
   * @param {EmailRow} row - The row data for the email to be sent
   */
  constructor(private row: EmailRow) {}

  /**
   * Sends an email based on the row data
   * @returns {boolean} True if email was sent successfully
   */
  public sendEmail(): boolean {
    // Get recipients
    const recipients = EmailUtils.combineEmailAddresses(
      this.row.distribution_emails,
      this.row.additional_emails
    );
    
    if (!recipients) {
      CustomLogger.debug('No valid email addresses found.');
      return false;
    }

    // Generate email body
    const emailBody = this.buildEmailBody();
    if (!emailBody) {
      CustomLogger.debug('Email body could not be generated.');
      return false;
    }

    // Get attachments
    const attachments = this.getAttachments();
    
    // Get final subject
    const subject = this.getFinalSubject();

    // Send email; if successful, trash the attachments.
    const sent = EmailUtils.sendEmail(recipients, subject, emailBody, attachments);
    if (sent) {
      this.trashAttachments();
    }
    return sent;
  }

  /**
   * Trashes all files specified in the attachments_urls column
   * This helps manage Drive storage by removing files that have already been sent
   * @private
   */
  private trashAttachments(): void {
    if (!this.row.attachments_urls) return;
    
    const fileUrls = this.row.attachments_urls.split(/[,;]+/).map(url => url.trim());
    for (const url of fileUrls) {
      try {
        const fileId = FileUtils.extractFileId(url);
        CustomLogger.debug(`Trashing file with fileId: ${fileId}`);
        FileUtils.trashFile(fileId);
      } catch (error) {
        // Get file metadata for better error logging
        let fileInfo: Record<string, any> = {};
        try {
          const fileId = FileUtils.extractFileId(url);
          fileInfo = FileUtils.getFileMetadata(fileId);
        } catch (subError) {
          fileInfo = { url, error: subError.toString() };
        }
        CustomLogger.error(`Error trashing file from URL ${url}`, { fileInfo, error });
      }
    }
  }

  /**
   * Generates the final email subject by processing the subject template
   * 
   * This method handles dynamic subject generation from templates. It works by:
   * 1. Taking the subject template from the row data
   * 2. Creating an Apps Script HTML template from the content
   * 3. Applying the template values to the template
   * 4. Evaluating the template to generate the final subject
   * 5. Decoding any HTML entities in the result
   * 
   * If any errors occur during this process, it falls back to the default subject
   * defined in the global configuration.
   * 
   * @returns {string} Final processed subject line for the email
   * @private
   */
  private getFinalSubject(): string {
    if (this.row.email_subject_template) {
      try {
        const subjectTemplate = HtmlService.createTemplate(this.row.email_subject_template);
        const values = this.getTemplateValues();
        Object.assign(subjectTemplate, values);
        
        let processedSubject = subjectTemplate.evaluate().getContent().trim();
        processedSubject = TextUtils.decodeHtmlEntities(processedSubject);
        return processedSubject || CONFIG.DEFAULT_SUBJECT;
      } catch (error) {
        CustomLogger.error('Error building subject from template', error);
        return CONFIG.DEFAULT_SUBJECT;
      }
    }
    return CONFIG.DEFAULT_SUBJECT;
  }

  /**
   * Creates the email body by applying template values to the HTML template
   * 
   * This method handles the dynamic generation of HTML email content. The process works as follows:
   * 1. Retrieves the HTML template content from Google Drive using the file URL in the row
   * 2. Gets all template values from the row's template_values JSON and any special fields
   * 3. Creates an Apps Script HTML template from the content
   * 4. Assigns all template values to the template object
   * 5. Evaluates the template, which processes all dynamic content markers (<?= varName ?>)
   *    and scriptlets (<? if (condition) { ?> content <? } ?>)
   * 
   * Note: The HTML template supports both simple variable substitution (<?= varName ?>)
   * and scriptlet conditionals/loops (<? if (condition) { ?> content <? } ?>)
   * 
   * @returns {string | null} Processed HTML email body or null if error occurs
   * @private
   * @see https://developers.google.com/apps-script/guides/html/templates
   */
  private buildEmailBody(): string | null {
    try {
      // Get template content and values
      const templateContent = FileUtils.getFileContentFromUrl(this.row.email_body_template);
      const values = this.getTemplateValues();
      
      // Apply template values to the HTML template
      const htmlTemplate = HtmlService.createTemplate(templateContent);
      
      // Assign all template values to the template
      Object.assign(htmlTemplate, values);
      
      // Evaluate the template - this processes all <?= varName ?> markers
      // and scriptlets (<? if (condition) { ?> content <? } ?>)
      return htmlTemplate.evaluate().getContent();
    } catch (error) {
      CustomLogger.error('Error building email body', error);
      return null;
    }
  }

  /**
   * Retrieves and returns the template values for the email body and subject
   * 
   * This method performs several important steps to prepare values for template processing:
   * 1. Sanitizes the JSON from template_values to handle common formatting issues
   * 2. Parses the sanitized JSON into a JavaScript object
   * 3. Extracts a Bluebeam session ID from revu_session_invite if available
   * 4. Merges all values into a single object that can be applied to templates
   * 
   * Special handling:
   * - The special 'sessionId' value is automatically extracted from revu_session_invite
   * - Template values are sanitized to handle malformed JSON (missing brackets, quotes, etc.)
   * - If JSON parsing fails, an empty object is used as a fallback
   * 
   * @returns {TemplateValues} Object containing all values to be used in templates
   * @private
   * @example
   * // If template_values contains: {"project":"Building A","pm":"John Doe"}
   * // And revu_session_invite contains: "Session ID: 123-456-789"
   * // Then getTemplateValues returns:
   * // {
   * //   "project": "Building A",
   * //   "pm": "John Doe",
   * //   "sessionId": "123-456-789"
   * // }
   */
  private getTemplateValues(): TemplateValues {
    // Sanitize JSON before parsing
    const sanitized = TextUtils.sanitizeJsonText(this.row.template_values);
    const values = sanitized ? JSON.parse(sanitized) : {};
    
    // Parse the session ID from revu_session_invite, if present
    const sessionId = EmailUtils.parseSessionId(this.row.revu_session_invite);
    if (sessionId) {
      values.sessionId = sessionId;
    }
    return values;
  }

  /**
   * Retrieves file attachments from Google Drive URLs
   * @returns {GoogleAppsScript.Base.Blob[]} Array of file blobs to attach
   * @private
   */
  private getAttachments(): GoogleAppsScript.Base.Blob[] {
    if (!this.row.attachments_urls) {
      CustomLogger.debug('No attachments provided.');
      return [];
    }

    const fileUrls = this.row.attachments_urls.split(/[,;]+/).map(url => url.trim());
    CustomLogger.debug(`Processing ${fileUrls.length} attachment(s).`);
    
    const attachments: GoogleAppsScript.Base.Blob[] = [];

    for (const url of fileUrls) {
      try {
        const fileId = FileUtils.extractFileId(url);
        CustomLogger.debug(`Retrieving attachment for fileId: ${fileId}`);
        
        // Check file size before attaching
        if (FileUtils.isFileTooLarge(fileId, MAX_ATTACHMENT_SIZE)) {
          const fileInfo = FileUtils.getFileMetadata(fileId);
          CustomLogger.error(`Attachment too large`, {
            fileId,
            size: fileInfo.size,
            maxSize: MAX_ATTACHMENT_SIZE
          });
          throw new Error(`Attachment file too large: ${fileId}`);
        }
        
        attachments.push(FileUtils.getFileBlob(fileId));
      } catch (error) {
        CustomLogger.error(`Error attaching file from URL ${url}`, error);
        throw error; // Re-throw to prevent sending email with missing attachments
      }
    }

    CustomLogger.debug(`Total attachments retrieved: ${attachments.length}`);
    return attachments;
  }
}

/**
 * Manages template data from the distribution_templates sheet
 * Used to populate default values for rows in the distribution_to_send sheet
 * 
 * The TemplateManager provides a template system that allows users to define reusable
 * email configurations in a separate sheet and apply them to multiple distributions.
 * This reduces duplication and ensures consistency across similar email distributions.
 */
class TemplateManager {
  private templatesSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private templateData: any[][] = [];
  private headerMap: { [key: string]: number } = {};
  private templateLabelIndex: { [label: string]: number } = {};

  /**
   * Creates a new TemplateManager instance
   * 
   * During initialization, the manager:
   * 1. Gets or creates the distribution_templates sheet
   * 2. Loads all template data from the sheet
   * 3. Creates an index of templates by their labels for fast lookup
   * 
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing template data
   */
  constructor(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    // Get or create the templates sheet
    this.templatesSheet = this.getOrCreateTemplatesSheet(spreadsheet);
    this.loadTemplateData();
  }

  /**
   * Gets or creates the distribution_templates sheet
   * 
   * This method ensures the templates sheet exists with appropriate headers.
   * If the sheet doesn't exist, it creates a new one with either:
   * - The same headers as the distributions_to_send sheet (if it exists)
   * - Default headers for essential fields (if the TO_SEND sheet doesn't exist)
   * 
   * In both cases, it ensures that 'template_label' column exists since this
   * is the key field used for template identification.
   * 
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to work with
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The templates sheet
   * @private
   */
  private getOrCreateTemplatesSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Spreadsheet.Sheet {
    // Use the SpreadsheetUtils helper to get or create the sheet
    let templatesSheet = spreadsheet.getSheetByName(SheetNames.TEMPLATES);
    if (!templatesSheet) {
      templatesSheet = spreadsheet.insertSheet(SheetNames.TEMPLATES);
      
      // Create the same headers as in the distributions_to_send sheet
      const toSendSheet = spreadsheet.getSheetByName(SheetNames.TO_SEND);
      if (toSendSheet) {
        const headers = toSendSheet.getRange(1, 1, 1, toSendSheet.getLastColumn()).getValues()[0];
        
        // Make sure template_label is in the headers
        if (!headers.includes('template_label')) {
          headers.push('template_label');
        }
        
        templatesSheet.appendRow(headers);
      } else {
        // Fallback headers if TO_SEND sheet doesn't exist
        templatesSheet.appendRow([
          'template_label', 'distribution_emails', 'additional_emails', 
          'revu_session_invite', 'template_values', 'email_body_template',
          'attachments_urls', 'email_subject_template', 'subject_template_value'
        ]);
      }
    }
    return templatesSheet;
  }

  /**
   * Loads template data from the templates sheet
   * 
   * This method:
   * 1. Retrieves all data from the templates sheet
   * 2. Maps column headers to their indices for easy reference
   * 3. Creates an index of template rows by their template_label value
   * 
   * The template label index provides O(1) lookup of templates by their label,
   * which is much more efficient than scanning the sheet each time.
   * 
   * @private
   */
  private loadTemplateData(): void {
    if (!this.templatesSheet) {
      CustomLogger.debug('Templates sheet not found, cannot load template data');
      return;
    }

    this.templateData = this.templatesSheet.getDataRange().getValues();
    if (this.templateData.length <= 1) {
      CustomLogger.debug('No template data available in the templates sheet');
      return;
    }

    // Map headers to column indices using the utility method
    const headers = this.templateData[0];
    this.headerMap = SpreadsheetUtils.mapHeadersToIndices(headers);

    // Create a template label index for quick lookups
    const templateLabelIndex = this.headerMap['template_label'];
    if (templateLabelIndex === undefined) {
      CustomLogger.debug('No template_label column found in templates sheet');
      return;
    }

    // Index all templates by their template label
    for (let i = 1; i < this.templateData.length; i++) {
      const label = this.templateData[i][templateLabelIndex];
      if (label) {
        this.templateLabelIndex[label] = i;
      }
    }

    CustomLogger.debug(`Loaded ${Object.keys(this.templateLabelIndex).length} templates from the templates sheet`);
  }

  /**
   * Gets template data for a specific template label
   * 
   * This method retrieves a template by its label and converts it to an EmailRow object.
   * It ensures all required fields are present in the returned object, even if they're
   * not defined in the template.
   * 
   * The template lookup process:
   * 1. Checks if the template label exists in the index
   * 2. If found, retrieves the corresponding row data
   * 3. Maps the row data to an EmailRow object with all required fields
   * 4. Sets the template_label explicitly to ensure it's included
   * 
   * @param {string} label - The template label to look up
   * @returns {EmailRow | null} The template data as an EmailRow object or null if not found
   * @public
   * @example
   * // Get a template called "RFI Response"
   * const template = templateManager.getTemplateByLabel("RFI Response");
   * if (template) {
   *   // Use the template values to populate an email row
   *   // Template fields will override any empty fields in the row
   * }
   */
  public getTemplateByLabel(label: string): EmailRow | null {
    if (!label || !this.templateLabelIndex[label]) {
      return null;
    }

    const rowIndex = this.templateLabelIndex[label];
    const templateRow = this.templateData[rowIndex];
    
    // Required fields that should be in every EmailRow object
    const requiredFields = [
      'template_label', 'distribution_emails', 'additional_emails',
      'revu_session_invite', 'template_values', 'email_body_template',
      'attachments_urls', 'email_subject_template', 'subject_template_value'
    ];
    
    // Convert the row to an EmailRow object using the SpreadsheetUtils helper
    const emailRow = SpreadsheetUtils.mapRowToObject(templateRow, this.headerMap, requiredFields) as EmailRow;
    
    // Set the template_label explicitly
    emailRow.template_label = label;
    
    return emailRow;
  }
}