/**
 * @fileoverview Email distribution system for review document management.
 * This script handles the automated sending of email distributions with attachments
 * using Google Apps Script. It processes data from a spreadsheet, sends emails with
 * customizable templates, and tracks sent distributions.
 * @version 1.2.1
 */

/** Defines the sheet names used in the spreadsheet */
enum SheetNames {
  TO_SEND       = 'distributions_to_send',
  SENT_HISTORY  = 'sent_history',
  TEMPLATES     = 'distribution_templates'
}

/** Global configuration for the email distribution system */
interface Config {
  readonly FROM_EMAIL: string;
  readonly DEFAULT_SUBJECT: string;
  readonly SPREADSHEET_ID: string;
}

/** Global configuration object */
const CONFIG: Config = {
  FROM_EMAIL:      'constdoc@ucsc.edu',
  DEFAULT_SUBJECT: 'Your Subject Here',
  SPREADSHEET_ID:  '1RfbiEpwU2APw3fXg5VoD4Dg2PlvacTV7EIrb_h98mcY'
};
const MAX_ATTACHMENT_SIZE = 21 * 1024 * 1024; // 21MB

/* ------------------------------------------------------------------ */
/*                              LOGGER                                */
/* ------------------------------------------------------------------ */
class AppScriptLogger {
  public static info(message: string, details: any = {}): void {
    console.log(`INFO  – ${message}`);
    Logger.log(`INFO  – ${message}`);
    if (details && Object.keys(details).length) {
      console.log(details);
      Logger.log(JSON.stringify(details, null, 2));
    }
  }
  public static warn(message: string, details: any = {}): void {
    console.warn(`WARN  – ${message}`);
    Logger.log(`WARN  – ${message}`);
    if (details && Object.keys(details).length) {
      console.warn(details);
      Logger.log(JSON.stringify(details, null, 2));
    }
  }
  public static error(message: string, error: any): void {
    console.error(`ERROR – ${message}`);
    Logger.log(`ERROR – ${message}`);
    console.error(error);
    Logger.log(error instanceof Error ? error.stack : JSON.stringify(error, null, 2));
  }
  /**
   * Formats an error message with a reference to the Apps Script execution logs.
   */
  public static formatErrorWithExecutionLogReference(message: string, error: any): string {
    return `${message}\n\nError details: ${error instanceof Error ? error.message : JSON.stringify(error)}\n\nFor complete details, check the Apps Script execution logs (View > Executions).`;
  }
}

/* ------------------------------------------------------------------ */
/*                         SPREADSHEET UTILS                          */
/* ------------------------------------------------------------------ */
class SpreadsheetUtils {
  /** Maps header names to zero-based column indices */
  static mapHeadersToIndices(headerRow: any[]): Record<string,number> {
    return headerRow.reduce((acc: Record<string,number>, h: string, i: number) => {
      acc[h] = i;
      return acc;
    }, {});
  }
  /** Gets or creates a sheet with optional headers */
  static getOrCreateSheet(
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
    name: string,
    headers?: string[]
  ): GoogleAppsScript.Spreadsheet.Sheet {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (headers && headers.length) {
        sheet.appendRow(headers);
      }
    }
    return sheet;
  }
  /**
   * Maps a row array to an object keyed by header names.
   * requiredFields ensures those keys exist even if empty.
   */
  static mapRowToObject(
    row: any[],
    headerMap: Record<string,number>,
    requiredFields: string[] = []
  ): Record<string,any> {
    const result: Record<string,any> = {};
    for (const f of requiredFields) {
      result[f] = headerMap[f] !== undefined ? row[headerMap[f]] || '' : '';
    }
    for (const h in headerMap) {
      if (!(h in result)) {
        result[h] = row[headerMap[h]] || '';
      }
    }
    return result;
  }
}

/* ------------------------------------------------------------------ */
/*                            FILE  UTILS                             */
/* ------------------------------------------------------------------ */
class FileUtils {
  static extractFileId(url: string): string {
    const m = url.match(/[-\w]{25,}/);
    if (!m) throw new Error(`Invalid Google Drive URL: ${url}`);
    return m[0];
  }
  static getFileContentFromUrl(url: string): string {
    const id = this.extractFileId(url);
    const blob = DriveApp.getFileById(id).getBlob();
    return blob.getDataAsString();
  }
  static isFileTooLarge(id: string, max: number): boolean {
    return DriveApp.getFileById(id).getSize() > max;
  }
  static getFileBlob(id: string): GoogleAppsScript.Base.Blob {
    return DriveApp.getFileById(id).getBlob();
  }
  static trashFile(id: string): boolean {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      return true;
    } catch(e) {
      AppScriptLogger.error(`Error trashing file ${id}`, e);
      return false;
    }
  }
  static getFileMetadata(id: string): Record<string,any> {
    try {
      const f = DriveApp.getFileById(id);
      return { id: f.getId(), name: f.getName(), size: f.getSize(), owner: f.getOwner()?.getEmail()||'' };
    } catch(e) {
      return { id, error: e.toString() };
    }
  }
}

/* ------------------------------------------------------------------ */
/*                          EMAIL  UTILS                             */
/* ------------------------------------------------------------------ */
class EmailUtils {
  static parseEmailAddresses(input: string): string[] {
    if (!input) return [];
    const clean = input.replace(/^\/\/.*/gm,'').toLowerCase();
    const re = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\s+at\s+)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\s+dot\s+))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;
    const out: string[] = [];
    let m: RegExpExecArray|null;
    while (m = re.exec(clean)) {
      let e = m[0].replace(/\s+(at|dot)\s+/g, (_,p) => p==='at'? '@':'.');
      out.push(e);
    }
    return out;
  }
  static parseSessionId(text: string): string|null {
    if (!text) return null;
    const m = text.match(/\b\d{3}-\d{3}-\d{3}\b/);
    return m ? m[0] : null;
  }
  static combineEmailAddresses(...sources: string[]): string|null {
    const all: string[] = [];
    for (const s of sources) {
      if (s) all.push(...this.parseEmailAddresses(s));
    }
    const uniq = Array.from(new Set(all.filter(Boolean)));
    return uniq.length ? uniq.join(',') : null;
  }
  static sendEmail(
    recipients: string,
    subject: string,
    htmlBody: string,
    attachments: GoogleAppsScript.Base.Blob[] = [],
    from: string = CONFIG.FROM_EMAIL
  ): boolean {
    try {
      GmailApp.sendEmail(recipients, subject, '', { htmlBody, attachments, from });
      return true;
    } catch(e) {
      AppScriptLogger.error('Error sending email', e);
      return false;
    }
  }
}

/* ------------------------------------------------------------------ */
/*                          TEXT  UTILS                              */
/* ------------------------------------------------------------------ */
class TextUtils {
  static decodeHtmlEntities(text: string): string {
    return text
      .replace(/&amp;/g,'&')
      .replace(/&lt;/g,'<')
      .replace(/&gt;/g,'>')
      .replace(/&quot;/g,'"')
      .replace(/&#39;/g,"'");
  }
  
  static sanitizeInput(text: string): string {
    if (!text) return '';
    // Replace ampersands with 'and' to prevent HTML entity encoding issues
    return text.replace(/&/g, 'and');
  }
  
  static sanitizeJsonText(text: string): string {
    if (!text) return '{}';
    try {
      let t = text.replace(/\\"/g,'"').replace(/\\\\/g,'\\').replace(/\r?\n/g,' ').trim();
      if (!t.startsWith('{')) t = '{'+t;
      if (!t.endsWith('}')) t += '}';
      const obj = JSON.parse(t);
      return JSON.stringify(obj);
    } catch(e) {
      AppScriptLogger.error('Error sanitizing JSON', e);
      return '{}';
    }
  }
}

/* ------------------------------------------------------------------ */
/*                         EMAIL  BUILDER                            */
/* ------------------------------------------------------------------ */
class EmailBuilder {
  constructor(private row: Record<string,any>) {
    // Sanitize subject template values to prevent HTML entity encoding issues
    if (this.row.subject_template_value) {
      this.row.subject_template_value = TextUtils.sanitizeInput(this.row.subject_template_value);
    }
    if (this.row.email_subject_template) {
      this.row.email_subject_template = TextUtils.sanitizeInput(this.row.email_subject_template);
    }
  }
  public sendEmail(): boolean {
    const to = EmailUtils.combineEmailAddresses(this.row.distribution_emails, this.row.additional_emails);
    if (!to) { AppScriptLogger.info('No recipients'); return false; }
    const body = this.buildEmailBody();
    if (!body) { AppScriptLogger.info('No body'); return false; }
    const atts = this.getAttachments();
    const subj = this.getFinalSubject();
    const sent = EmailUtils.sendEmail(to, subj, body, atts);
    if (sent) this.trashAttachments();
    return sent;
  }
  private buildEmailBody(): string|null {
    try {
      const html = FileUtils.getFileContentFromUrl(this.row.email_body_template);
      const tpl = HtmlService.createTemplate(html);
      Object.assign(tpl, this.getTemplateValues());
      return tpl.evaluate().getContent();
    } catch(e) {
      AppScriptLogger.error('Error building body', e);
      return null;
    }
  }
  private getFinalSubject(): string {
    if (!this.row.email_subject_template) return CONFIG.DEFAULT_SUBJECT;
    try {
      const tpl = HtmlService.createTemplate(this.row.email_subject_template);
      Object.assign(tpl, this.getTemplateValues());
      let s = tpl.evaluate().getContent().trim();
      return TextUtils.decodeHtmlEntities(s) || CONFIG.DEFAULT_SUBJECT;
    } catch(e) {
      AppScriptLogger.error('Error building subject', e);
      return CONFIG.DEFAULT_SUBJECT;
    }
  }
  private getTemplateValues(): Record<string,any> {
    const json = TextUtils.sanitizeJsonText(this.row.email_template_values);
    const vals = JSON.parse(json);
    const sid = EmailUtils.parseSessionId(this.row.revu_session_invite);
    if (sid) vals.sessionId = sid;
    return vals;
  }
  private getAttachments(): GoogleAppsScript.Base.Blob[] {
    if (!this.row.attachments_urls) return [];
    const blobs: GoogleAppsScript.Base.Blob[] = [];
    for (const url of this.row.attachments_urls.split(/[,;]+/).map(s=>s.trim())) {
      try {
        const id = FileUtils.extractFileId(url);
        if (FileUtils.isFileTooLarge(id, MAX_ATTACHMENT_SIZE)) {
          throw new Error(`Attachment too large: ${id}`);
        }
        blobs.push(FileUtils.getFileBlob(id));
      } catch(e) {
        AppScriptLogger.error(`Error attaching ${url}`, e);
        throw e;
      }
    }
    return blobs;
  }
  private trashAttachments(): void {
    if (!this.row.attachments_urls) return;
    for (const url of this.row.attachments_urls.split(/[,;]+/).map(s=>s.trim())) {
      try {
        FileUtils.trashFile(FileUtils.extractFileId(url));
      } catch {}
    }
  }
}

/* ------------------------------------------------------------------ */
/*                       TEMPLATE  MANAGEMENT                         */
/* ------------------------------------------------------------------ */
class TemplateManager {
  private templatesSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private templateData: any[][] = [];
  private headerMap: Record<string,number> = {};
  private index: Record<string,number> = {};

  constructor(private ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    this.templatesSheet = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TEMPLATES);
    this.loadTemplateData();
  }
  private loadTemplateData(): void {
    const rows = this.templatesSheet.getDataRange().getValues();
    if (rows.length < 2) return;
    this.headerMap = SpreadsheetUtils.mapHeadersToIndices(rows[0]);
    for (let i = 1; i < rows.length; i++) {
      const label = rows[i][this.headerMap['distribution_template_label']];
      if (label) {
        this.index[label] = i;
        this.templateData[i] = rows[i];
      }
    }
  }
  public getTemplateByLabel(label: string): Record<string,any>|null {
    const idx = this.index[label];
    if (idx == null) return null;
    return SpreadsheetUtils.mapRowToObject(this.templateData[idx], this.headerMap, Object.keys(this.headerMap));
  }
  public getAvailableTemplates(): string[] {
    return Object.keys(this.index);
  }
}

/* ------------------------------------------------------------------ */
/*                         EMAIL  PROCESSOR                           */
/* ------------------------------------------------------------------ */
/** EmailProcessor handles the main logic for sending emails and applying templates */
class EmailProcessor {
  private ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  private source = this.ss.getSheetByName(SheetNames.TO_SEND)!;
  private history = SpreadsheetUtils.getOrCreateSheet(this.ss, SheetNames.SENT_HISTORY, [...this.getHeaderNames(), 'datetime']);
  private tm = new TemplateManager(this.ss);

  private getHeaderNames(): string[] {
    return this.source.getRange(1,1,1,this.source.getLastColumn()).getValues()[0].filter(Boolean) as string[];
  }
  private getHeaderMap(): Record<string,number> {
    return SpreadsheetUtils.mapHeadersToIndices(
      this.source.getRange(1,1,1,this.source.getLastColumn()).getValues()[0]
    );
  }

  /** sendEmails processes and moves rows to history */
  public sendEmails(): void {
    const data = this.source.getDataRange().getValues();
    const hMap = this.getHeaderMap();
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row.every(c=>!c)) continue;
      let emailRow: Record<string,any>;
      try {
        emailRow = SpreadsheetUtils.mapRowToObject(row, hMap) as Record<string,any>;
        // apply template
        if (emailRow.distribution_template_label) {
          const tpl = this.tm.getTemplateByLabel(emailRow.distribution_template_label);
          if (!tpl) throw new Error(`Template "${emailRow.distribution_template_label}" not found`);
          for (const k in tpl) {
            if (!emailRow[k] && tpl[k]) emailRow[k] = tpl[k];
          }
        }
      } catch(e) {
        AppScriptLogger.error(`Template error row ${i+1}`, e);
        SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference(`Error with template in row ${i+1}`, e));
        continue;
      }
      try {
        const builder = new EmailBuilder(emailRow);
        if (builder.sendEmail()) {
          this.history.appendRow([...row, new Date()]);
          this.source.deleteRow(i+1);
        }
      } catch(e) {
        AppScriptLogger.error(`Error processing row ${i+1}`, e);
      }
    }
  }

  /** applyTemplatesToPendingRows collects errors and updates empty cells */
  public applyTemplatesToPendingRows(): { updatedRowCount:number; errors:string[] } {
    const data = this.source.getDataRange().getValues();
    const hMap = this.getHeaderMap();
    let updated = 0;
    const errors: string[] = [];

    for (let i = 1; i < data.length; i++) {
      const sheetRow = data[i];
      const rowIndex = i + 1;
      const er = SpreadsheetUtils.mapRowToObject(sheetRow, hMap) as Record<string,any>;
      if (!er.distribution_template_label) continue;

      const tpl = this.tm.getTemplateByLabel(er.distribution_template_label);
      if (!tpl) {
        errors.push(`Row ${rowIndex}: template "${er.distribution_template_label}" not found`);
        continue;
      }
      let changed = false;
      for (const k in tpl) {
        const col = hMap[k];
        if (col == null) continue;
        const val = sheetRow[col];
        if ((!val || val === '') && tpl[k]) {
          this.source.getRange(rowIndex, col+1).setValue(tpl[k]);
          changed = true;
        }
      }
      if (changed) updated++;
    }
    return { updatedRowCount: updated, errors };
  }
}

/* ------------------------------------------------------------------ */
/*                         GLOBAL  FUNCTIONS                          */
/* ------------------------------------------------------------------ */

/** Applies templates to pending rows and shows a consolidated alert */
function applyTemplatesToPendingRows(): void {
  try {
    AppScriptLogger.info('Applying templates to pending rows...');
    const res = new EmailProcessor().applyTemplatesToPendingRows();
    AppScriptLogger.info(`Template application completed. Updated ${res.updatedRowCount} rows.`);
    const ui = SpreadsheetApp.getUi();
    if (res.errors.length) {
      const msg = AppScriptLogger.formatErrorWithExecutionLogReference(
        `Some errors occurred while applying templates:\n\n${res.errors.join('\n')}`, {}
      );
      ui.alert(msg);
    }
  } catch(e) {
    AppScriptLogger.error('Error applying templates', e);
    SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error applying templates', e));
  }
}

/** Processes all pending email distributions */
function processEmailDistributions(): void {
  try {
    AppScriptLogger.info('Processing email distributions...');
    new EmailProcessor().sendEmails();
    AppScriptLogger.info('Email distribution completed.');
  } catch(e) {
    AppScriptLogger.error('Error processing email distributions', e);
    SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error processing email distributions', e));
  }
}

/** Initializes spreadsheet structure */
function initializeSpreadsheetStructure(): void {
  try {
    AppScriptLogger.info('Initializing spreadsheet structure...');
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const std = [
      'distribution_template_label','distribution_emails','additional_emails',
      'revu_session_invite','email_template_values','email_body_template',
      'attachments_urls','email_subject_template','subject_template_value'
    ];
    const toSend = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TO_SEND, std);
    const hist   = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.SENT_HISTORY, [...std,'datetime']);
    const tmpl   = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TEMPLATES, std);
    [toSend,hist,tmpl].forEach(sh => {
      if (sh.getLastRow()>0) {
        sh.getRange(1,1,1,sh.getLastColumn())
          .setBackground('#f3f3f3')
          .setFontWeight('bold');
      }
    });
    SpreadsheetApp.getUi().alert('Spreadsheet structure initialized successfully.');
  } catch(e) {
    AppScriptLogger.error('Error initializing spreadsheet structure', e);
    SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error initializing spreadsheet structure', e));
  }
}

/** Builds the custom menu on open */
function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Email Distributions')
    .addItem('Send Pending Emails','processEmailDistributions')
    .addItem('Apply Templates to Pending Rows','applyTemplatesToPendingRows')
    .addItem('Initialize Spreadsheet Structure','initializeSpreadsheetStructure')
    .addToUi();
}