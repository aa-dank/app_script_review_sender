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
  SENT_HISTORY = 'sent_history'
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
  FROM_EMAIL: 'constdoc@gmail.com',
  DEFAULT_SUBJECT: 'Your Subject Here',
  SPREADSHEET_ID: '1RfbiEpwU2APw3fXg5VoD4Dg2PlvacTV7EIrb_h98mcY',
  DEBUG_MODE: true
};

// Types and Interfaces
/**
 * Represents a row of email distribution data from the spreadsheet
 * @interface EmailRow
 */
interface EmailRow {
  /** Primary distribution list email addresses */
  distribution_email: string;
  /** Additional individual email addresses to include */
  additional_emails: string;
  /** Bluebeam Revu session invite text containing session ID */
  revu_session_invite: string;
  /** JSON string containing template variable values */
  body_template_values: string;
  /** Google Drive URL of the email template HTML file */
  email_body_template: string;
  /** Comma-separated list of Google Drive URLs for attachments */
  files: string;
  /** Custom subject line for the email */
  email_subject: string;
  email_subject_template: string;
  subject_template_values: string;
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
 * Main processor class for handling email distributions
 * Manages spreadsheet operations and coordinates email sending process
 */
class EmailProcessor {
  private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private sourceSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private historySheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    this.spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    this.sourceSheet = this.spreadsheet.getSheetByName(SheetNames.TO_SEND);
    this.historySheet = this.getOrCreateHistorySheet();
  }

  /**
   * Creates or retrieves the history sheet for tracking sent distributions
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The history sheet
   * @private
   */
  private getOrCreateHistorySheet(): GoogleAppsScript.Spreadsheet.Sheet {
    let historySheet = this.spreadsheet.getSheetByName(SheetNames.SENT_HISTORY);
    if (!historySheet) {
      historySheet = this.spreadsheet.insertSheet(SheetNames.SENT_HISTORY);
      const headers = [
        ...Object.keys(this.getHeaders()),
        'datetime'
      ];
      historySheet.appendRow(headers);
    }
    return historySheet;
  }

  private getHeaders(): { [key: string]: number } {
    const headers = this.sourceSheet.getRange(1, 1, 1, this.sourceSheet.getLastColumn()).getValues()[0];
    return headers.reduce((acc: { [key: string]: number }, header: string, index: number) => {
      acc[header] = index;
      return acc;
    }, {});
  }

  /**
   * Processes all pending email distributions in the spreadsheet
   * @returns {Promise<void>}
   * @public
   */
  public async sendEmails(): Promise<void> {
    const data = this.sourceSheet.getDataRange().getValues();
    const headers = this.getHeaders();

    // Process rows in reverse order
    for (let i = data.length - 1; i >= 1; i--) {
      const row = this.mapRowToEmailRow(data[i], headers);
      
      try {
        if (await this.processRow(row)) {
          this.moveRowToHistory(i + 1, data[i]);
        }
      } catch (error) {
        CustomLogger.error(`Error processing row ${i + 1}`, error);
      }
    }
  }

  /**
   * Maps a raw data row to a typed EmailRow object
   * @param {any[]} row - Raw spreadsheet row data
   * @param {{ [key: string]: number }} headers - Map of header names to column indices
   * @returns {EmailRow} Typed email row object
   * @private
   */
  private mapRowToEmailRow(row: any[], headers: { [key: string]: number }): EmailRow {
    return {
      distribution_email: row[headers['distribution_email']] || '',
      additional_emails: row[headers['additional_emails']] || '',
      revu_session_invite: row[headers['revu_session_invite']] || '',
      body_template_values: row[headers['body_template_values']] || '',
      email_body_template: row[headers['email_body_template']] || '',
      files: row[headers['files']] || '',
      email_subject: row[headers['email_subject']] || CONFIG.DEFAULT_SUBJECT,
      email_subject_template: row[headers['email_subject_template']] || '',
      subject_template_values: row[headers['subject_template_values']] || ''
    };
  }

  private async processRow(row: EmailRow): Promise<boolean> {
    if (!this.validateRow(row)) {
      return false;
    }

    const emailProcessor = new EmailBuilder(row);
    return await emailProcessor.sendEmail();
  }

  private validateRow(row: EmailRow): boolean {
    if (!row.distribution_email && !row.additional_emails) {
      CustomLogger.debug('No recipient email addresses provided.');
      return false;
    }
    if (!row.email_body_template) {
      CustomLogger.debug('No email template URL provided.');
      return false;
    }
    return true;
  }

  private moveRowToHistory(rowIndex: number, rowData: any[]): void {
    const rowWithTimestamp = [...rowData, new Date()];
    this.historySheet.appendRow(rowWithTimestamp);
    this.sourceSheet.deleteRow(rowIndex);
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

  public async sendEmail(): Promise<boolean> {
    const recipients = this.compileEmailAddresses();
    if (!recipients) return false;

    const emailBody = await this.buildEmailBody();
    if (!emailBody) return false;

    const attachments = await this.getAttachments();
    const subject = this.getFinalSubject();

    return this.sendEmailViaGmail(recipients, emailBody, attachments, subject);
  }

  private getFinalSubject(): string {
    if (this.row.email_subject_template) {
      try {
        const templateValues = this.row.subject_template_values
          ? JSON.parse(this.row.subject_template_values)
          : {};
        const subjectTemplate = HtmlService.createTemplate(this.row.email_subject_template);
        Object.assign(subjectTemplate, templateValues);
        return subjectTemplate.evaluate().getContent().trim() || CONFIG.DEFAULT_SUBJECT;
      } catch (error) {
        CustomLogger.error('Error building subject from template', error);
      }
    }
    // Return existing row subject or fallback to DEFAULT_SUBJECT
    return this.row.email_subject || CONFIG.DEFAULT_SUBJECT;
  }

  /**
   * Compiles and validates email addresses from distribution and additional emails
   * @returns {string | null} Comma-separated list of valid email addresses or null if none found
   * @private
   */
  private compileEmailAddresses(): string | null {
    const emails = [
      ...this.parseEmails(this.row.distribution_email),
      ...this.parseEmails(this.row.additional_emails)
    ];

    const uniqueEmails = [...new Set(emails.filter(Boolean))];
    return uniqueEmails.length > 0 ? uniqueEmails.join(',') : null;
  }

  /**
   * Extracts email addresses from a string, handling various formats including 'at' and 'dot' notation
   * @param {string} input - Raw string containing email addresses
   * @returns {string[]} Array of valid email addresses
   * @private
   */
  private parseEmails(input: string): string[] {
    if (!input) return [];
    
    const cleanInput = input.replace(/^\/\/.*/gm, '').toLowerCase();
    const emailRegex = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\s+at\s+)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\s+dot\s+))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;

    const matches: string[] = [];
    let match: RegExpExecArray | null;
    
    while ((match = emailRegex.exec(cleanInput)) !== null) {
      if (!match[0].startsWith('//')) {
        const email = match[0].replace(/\s+(at|dot)\s+/g, (m, p1) => 
          p1 === 'at' ? '@' : '.'
        );
        matches.push(email);
      }
    }

    return matches;
  }

  /**
   * Creates the email body by applying template values to the HTML template
   * @returns {Promise<string | null>} Processed HTML email body or null if error occurs
   * @private
   */
  private async buildEmailBody(): Promise<string | null> {
    try {
      const templateContent = await this.getTemplateContent();
      const values = this.getTemplateValues();
      
      const htmlTemplate = HtmlService.createTemplate(templateContent);
      Object.assign(htmlTemplate, values);
      
      return htmlTemplate.evaluate().getContent();
    } catch (error) {
      CustomLogger.error('Error building email body', error);
      return null;
    }
  }

  private async getTemplateContent(): Promise<string> {
    const fileId = this.extractFileId(this.row.email_body_template);
    const templateFile = DriveApp.getFileById(fileId);
    return templateFile.getBlob().getDataAsString();
  }

  private getTemplateValues(): TemplateValues {
    const values: TemplateValues = this.row.body_template_values ? 
      JSON.parse(this.row.body_template_values) : {};

    const sessionId = this.parseSessionId(this.row.revu_session_invite);
    if (sessionId) {
      values.sessionId = sessionId;
    }

    return values;
  }

  /**
   * Extracts a Bluebeam session ID from the provided text
   * @param {string} text - Text containing the session ID
   * @returns {string | null} Extracted session ID or null if not found
   * @private
   */
  private parseSessionId(text: string): string | null {
    if (!text) return null;

    const matches = text.match(/\D(\d{3}-\d{3}-\d{3})(?=$|\D)/g);
    if (!matches) return null;

    const ids = matches.map(m => m.match(/\d{3}-\d{3}-\d{3}/)?.[0]).filter(Boolean);
    
    if (ids.length === 0) {
      CustomLogger.debug("WARNING: No Bluebeam Session ID found.");
      return null;
    }

    if (new Set(ids).size > 1) {
      CustomLogger.debug("WARNING: Multiple different session IDs found. Using first match.");
    }

    return ids[0] || null;
  }

  /**
   * Retrieves file attachments from Google Drive URLs
   * @returns {Promise<GoogleAppsScript.Base.Blob[]>} Array of file blobs to attach
   * @private
   */
  private async getAttachments(): Promise<GoogleAppsScript.Base.Blob[]> {
    if (!this.row.files) return [];

    const fileUrls = this.row.files.split(/[,;]+/).map(url => url.trim());
    const attachments: GoogleAppsScript.Base.Blob[] = [];

    for (const url of fileUrls) {
      try {
        const fileId = this.extractFileId(url);
        const file = DriveApp.getFileById(fileId);
        attachments.push(file.getBlob());
      } catch (error) {
        CustomLogger.error(`Error attaching file from URL ${url}`, error);
      }
    }

    return attachments;
  }

  private extractFileId(url: string): string {
    const match = url.match(/[-\w]{25,}/);
    if (!match?.[0]) {
      throw new Error(`Invalid Google Drive URL: ${url}`);
    }
    return match[0];
  }

  /**
   * Sends the email using Gmail service
   * @param {string} recipients - Comma-separated list of recipient email addresses
   * @param {string} htmlBody - HTML content of the email
   * @param {GoogleAppsScript.Base.Blob[]} attachments - Array of file attachments
   * @param {string} subject - Subject of the email
   * @returns {boolean} True if email was sent successfully
   * @private
   */
  private sendEmailViaGmail(
    recipients: string, 
    htmlBody: string, 
    attachments: GoogleAppsScript.Base.Blob[],
    subject: string
  ): boolean {
    try {
      GmailApp.sendEmail(recipients, subject, '', {
        htmlBody,
        attachments,
        from: CONFIG.FROM_EMAIL
      });
      return true;
    } catch (error) {
      CustomLogger.error('Error sending email', error);
      return false;
    }
  }
}

/**
 * Enhanced logging utility
 */
class CustomLogger {
  static debug(message: string, data?: any) {
      if (CONFIG.DEBUG_MODE) {
          Logger.log(`[DEBUG] ${message} ${data ? JSON.stringify(data) : ''}`);
      }
  }

  static error(message: string, error: any) {
      Logger.log(`[ERROR] ${message}`);
      Logger.log('Error details:' + JSON.stringify({
          message: error.message,
          name: error.name,
          stack: error.stack,
          toString: error.toString()
      }));
  }
}

/**
 * Main entrypoint function that can be called from Apps Script console.
 * This function processes all pending email distributions in the spreadsheet.
 */
async function processEmailDistributions(): Promise<void> {
  try {
      CustomLogger.debug('Doin it.')  
      const processor = new EmailProcessor();
      await processor.sendEmails();
      CustomLogger.debug('Email distribution processing completed.');
  } catch (error) {
      CustomLogger.error('Error processing email distributions.', error);
  }
}

// Add this if you want to create a custom menu in the spreadsheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Distributions')
    .addItem('Send Pending Emails', 'processEmailDistributions')
    .addToUi();
}

/**
 * Validates spreadsheet access and returns the spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 * @throws {Error} If spreadsheet cannot be accessed
 */
function getSpreadsheet() {
  try {
    CustomLogger.debug('Attempting to access spreadsheet', {
      spreadsheetId: CONFIG.SPREADSHEET_ID
    });

    // Validate spreadsheet ID
    if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'your-spreadsheet-id') {
      throw new Error('Invalid spreadsheet ID configuration');
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    if (!spreadsheet) {
      throw new Error('Could not open spreadsheet - null response');
    }

    CustomLogger.debug('Successfully accessed spreadsheet');
    return spreadsheet;

  } catch (error) {
    CustomLogger.error('Spreadsheet access failed', error);
    throw new Error(`Spreadsheet access error: ${error.message || error}`);
  }
}

// Add this test function to verify access
function testSpreadsheetAccess() {
  try {
    CustomLogger.debug('Starting spreadsheet access test');
    
    // Check if SpreadsheetApp is available
    CustomLogger.debug('Checking SpreadsheetApp availability', {
      hasSpreadsheetApp: typeof SpreadsheetApp !== 'undefined'
    });

    const spreadsheet = getSpreadsheet();
    CustomLogger.debug('Spreadsheet accessed successfully', {
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl()
    });

  } catch (error) {
    CustomLogger.error('Access test failed', error);
  }
}

function testSpreadsheetPermissions() {
  CustomLogger.debug('=== Starting Permission Tests ===');
  
  try {
    // Test 1: Basic SpreadsheetApp access
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    CustomLogger.debug('Active spreadsheet test:', {
      name: activeSpreadsheet.getName(),
      id: activeSpreadsheet.getId()
    });
    
    // Test 2: Create temporary test spreadsheet
    const testSheet = SpreadsheetApp.create('Test Sheet');
    const testId = testSheet.getId();
    CustomLogger.debug('Created test spreadsheet:', { id: testId });
    
    // Test 3: Open by ID
    const openedSheet = SpreadsheetApp.openById(testId);
    CustomLogger.debug('Opened test spreadsheet by ID');
    
    // Cleanup
    DriveApp.getFileById(testId).setTrashed(true);
    CustomLogger.debug('Test spreadsheet deleted');
    
    return true;
  } catch (error) {
    CustomLogger.error('Permission test failed at step:', error);
    return false;
  }
}

/**
 * Verify services are enabled
 */
function checkServices() {
  CustomLogger.debug('=== Checking Services ===');
  
  try {
    const services = {
      spreadsheet: typeof SpreadsheetApp !== 'undefined',
      drive: typeof DriveApp !== 'undefined'
    };
    
    CustomLogger.debug('Service availability:', services);
    return services;
  } catch (error) {
    CustomLogger.error('Service check failed:', error);
    return false;
  }
}