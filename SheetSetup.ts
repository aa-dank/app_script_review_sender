/**
 * Global configuration interface for the email distribution system
 * @interface Config
 */
export interface Config {
  /** Email address to send from (requires Gmail "Send As" permissions) */
  readonly FROM_EMAIL: string;
  /** Default subject line for emails when none is provided */
  readonly DEFAULT_SUBJECT: string;
  /** ID of the Google Spreadsheet containing distribution data */
  readonly SPREADSHEET_ID: string;
  /** Enable debug logging */
  readonly DEBUG_MODE: boolean;
}

/**
 * Represents a row of email distribution data from the spreadsheet
 * @interface EmailRow
 */
export interface EmailRow {
  /** Primary distribution list email addresses */
  distribution_emails: string;
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
  subject_template_value: string;
  /** Allow for additional dynamic columns */
  [key: string]: string; // Allow additional columns
}

export enum SheetNames {
    TO_SEND = 'distributions_to_send',
    SENT_HISTORY = 'sent_history'
  }
  
export const CONFIG: Config = {
  FROM_EMAIL: 'constdoc@ucsc.edu',
  DEFAULT_SUBJECT: 'Your Subject Here',
  SPREADSHEET_ID: '1RfbiEpwU2APw3fXg5VoD4Dg2PlvacTV7EIrb_h98mcY',
  DEBUG_MODE: true
};
  
export function mapRowToEmailRow(
  row: any[],
  headers: { [key: string]: number }
): EmailRow {
  return {
    distribution_emails: row[headers['distribution_emails']] || '',
    additional_emails: row[headers['additional_emails']] || '',
    revu_session_invite: row[headers['revu_session_invite']] || '',
    body_template_values: row[headers['body_template_values']] || '',
    email_body_template: row[headers['email_body_template']] || '',
    files: row[headers['files']] || '',
    email_subject: row[headers['email_subject']] || CONFIG.DEFAULT_SUBJECT,
    email_subject_template: row[headers['email_subject_template']] || '',
    subject_template_value: row[headers['subject_template_value']] || ''
  };
}