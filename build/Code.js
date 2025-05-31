/**
 * @fileoverview Email distribution system for review document management.
 * This script handles the automated sending of email distributions with attachments
 * using Google Apps Script. It processes data from a spreadsheet, sends emails with
 * customizable templates, and tracks sent distributions.
 * @version 1.2.3
 */
/** Defines the sheet names used in the spreadsheet */
var SheetNames;
(function (SheetNames) {
    SheetNames["TO_SEND"] = "distributions_to_send";
    SheetNames["SENT_HISTORY"] = "sent_history";
    SheetNames["TEMPLATES"] = "distribution_templates";
})(SheetNames || (SheetNames = {}));
/** Global configuration object */
const CONFIG = {
    FROM_EMAIL: 'constdoc@ucsc.edu',
    SPREADSHEET_ID: '1RfbiEpwU2APw3fXg5VoD4Dg2PlvacTV7EIrb_h98mcY'
};
const MAX_ATTACHMENT_SIZE = 21 * 1024 * 1024; // 21MB
/* ------------------------------------------------------------------ */
/*                              LOGGER                                */
/* ------------------------------------------------------------------ */
class AppScriptLogger {
    static info(message, details = {}) {
        console.log(`INFO  – ${message}`);
        Logger.log(`INFO  – ${message}`);
        if (details && Object.keys(details).length) {
            console.log(details);
            Logger.log(JSON.stringify(details, null, 2));
        }
    }
    static warn(message, details = {}) {
        console.warn(`WARN  – ${message}`);
        Logger.log(`WARN  – ${message}`);
        if (details && Object.keys(details).length) {
            console.warn(details);
            Logger.log(JSON.stringify(details, null, 2));
        }
    }
    static error(message, error) {
        console.error(`ERROR – ${message}`);
        Logger.log(`ERROR – ${message}`);
        console.error(error);
        Logger.log(error instanceof Error ? error.stack : JSON.stringify(error, null, 2));
    }
    /**
     * Formats an error message with a reference to the Apps Script execution logs.
     */
    static formatErrorWithExecutionLogReference(message, error) {
        return `${message}\n\nError details: ${error instanceof Error ? error.message : JSON.stringify(error)}\n\nFor complete details, check the Apps Script execution logs (View > Executions).`;
    }
}
/* ------------------------------------------------------------------ */
/*                         SPREADSHEET UTILS                          */
/* ------------------------------------------------------------------ */
class SpreadsheetUtils {
    /** Maps header names to zero-based column indices */
    static mapHeadersToIndices(headerRow) {
        return headerRow.reduce((acc, h, i) => {
            acc[h] = i;
            return acc;
        }, {});
    }
    /** Gets or creates a sheet with optional headers */
    static getOrCreateSheet(ss, name, headers) {
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
    static mapRowToObject(row, headerMap, requiredFields = []) {
        const result = {};
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
    static extractFileId(url) {
        const m = url.match(/[-\w]{25,}/);
        if (!m)
            throw new Error(`Invalid Google Drive URL: ${url}`);
        return m[0];
    }
    static getFileContentFromUrl(url) {
        const id = this.extractFileId(url);
        const blob = DriveApp.getFileById(id).getBlob();
        return blob.getDataAsString();
    }
    static isFileTooLarge(id, max) {
        return DriveApp.getFileById(id).getSize() > max;
    }
    static getFileBlob(id) {
        return DriveApp.getFileById(id).getBlob();
    }
    static trashFile(id) {
        try {
            DriveApp.getFileById(id).setTrashed(true);
            return true;
        }
        catch (e) {
            AppScriptLogger.error(`Error trashing file ${id}`, e);
            return false;
        }
    }
    static getFileMetadata(id) {
        var _a;
        try {
            const f = DriveApp.getFileById(id);
            return { id: f.getId(), name: f.getName(), size: f.getSize(), owner: ((_a = f.getOwner()) === null || _a === void 0 ? void 0 : _a.getEmail()) || '' };
        }
        catch (e) {
            return { id, error: e.toString() };
        }
    }
}
/* ------------------------------------------------------------------ */
/*                          EMAIL  UTILS                             */
/* ------------------------------------------------------------------ */
class EmailUtils {
    static parseEmailAddresses(input) {
        if (!input)
            return [];
        const clean = input.replace(/^\/\/.*/gm, '').toLowerCase();
        const re = /([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*(@|\s+at\s+)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|\s+dot\s+))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)/gi;
        const out = [];
        let m;
        while (m = re.exec(clean)) {
            let e = m[0].replace(/\s+(at|dot)\s+/g, (_, p) => p === 'at' ? '@' : '.');
            out.push(e);
        }
        return out;
    }
    static parseSessionId(text) {
        if (!text)
            return null;
        const m = text.match(/\b\d{3}-\d{3}-\d{3}\b/);
        return m ? m[0] : null;
    }
    static combineEmailAddresses(...sources) {
        const all = [];
        for (const s of sources) {
            if (s)
                all.push(...this.parseEmailAddresses(s));
        }
        const uniq = Array.from(new Set(all.filter(Boolean)));
        return uniq.length ? uniq.join(',') : null;
    }
    static sendEmail(recipients, subject, htmlBody, attachments = [], from = CONFIG.FROM_EMAIL) {
        try {
            GmailApp.sendEmail(recipients, subject, '', { htmlBody, attachments, from });
            return true;
        }
        catch (e) {
            AppScriptLogger.error('Error sending email', e);
            return false;
        }
    }
}
/* ------------------------------------------------------------------ */
/*                          TEXT  UTILS                              */
/* ------------------------------------------------------------------ */
class TextUtils {
    static decodeHtmlEntities(text) {
        return text
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'");
    }
    static sanitizeInput(text) {
        if (!text)
            return '';
        // Replace ampersands with 'and' to prevent HTML entity encoding issues
        return text.replace(/&/g, 'and');
    }
    static sanitizeJsonText(text) {
        if (!text)
            return '{}';
        try {
            let t = text.replace(/\\"/g, '"').replace(/\\\\/g, '\\').replace(/\r?\n/g, ' ').trim();
            if (!t.startsWith('{'))
                t = '{' + t;
            if (!t.endsWith('}'))
                t += '}';
            const obj = JSON.parse(t);
            return JSON.stringify(obj);
        }
        catch (e) {
            AppScriptLogger.error('Error sanitizing JSON', e);
            return '{}';
        }
    }
}
/* ------------------------------------------------------------------ */
/*                         EMAIL  BUILDER                            */
/* ------------------------------------------------------------------ */
class EmailBuilder {
    constructor(row) {
        this.row = row;
        // Sanitize subject template values to prevent HTML entity encoding issues
        if (this.row.subject_template_value) {
            this.row.subject_template_value = TextUtils.sanitizeInput(this.row.subject_template_value);
        }
        if (this.row.email_subject_template) {
            this.row.email_subject_template = TextUtils.sanitizeInput(this.row.email_subject_template);
        }
    }
    sendEmail() {
        const to = EmailUtils.combineEmailAddresses(this.row.distribution_emails, this.row.additional_emails);
        if (!to) {
            AppScriptLogger.info('No recipients');
            return false;
        }
        const body = this.buildEmailBody();
        if (!body) {
            AppScriptLogger.info('No body');
            return false;
        }
        let subj;
        try {
            subj = this.getFinalSubject();
            if (!subj) {
                AppScriptLogger.info('Empty subject line');
                return false;
            }
        }
        catch (e) {
            AppScriptLogger.info('Failed to generate subject line');
            return false;
        }
        const atts = this.getAttachments();
        const sent = EmailUtils.sendEmail(to, subj, body, atts);
        if (sent)
            this.trashAttachments();
        return sent;
    }
    buildEmailBody() {
        try {
            const html = FileUtils.getFileContentFromUrl(this.row.email_body_template);
            // Find all template variables in the HTML (looking for <%= varName %> patterns)
            const templateVarRegex = /<%=\s*([a-zA-Z0-9_]+)\s*%>/g;
            const templateVars = new Set();
            let match;
            while ((match = templateVarRegex.exec(html)) !== null) {
                templateVars.add(match[1]);
            }
            // Get template values and ensure all referenced variables exist
            const templateValues = this.getTemplateValues();
            for (const varName of templateVars) {
                if (templateValues[varName] === undefined) {
                    AppScriptLogger.warn(`Template uses undefined variable: ${varName}. Using empty string.`, {
                        rowData: JSON.stringify(this.row)
                    });
                    templateValues[varName] = '';
                }
            }
            const tpl = HtmlService.createTemplate(html);
            Object.assign(tpl, templateValues);
            return tpl.evaluate().getContent();
        }
        catch (e) {
            AppScriptLogger.error('Error building body', e);
            return null;
        }
    }
    getFinalSubject() {
        if (!this.row.email_subject_template) {
            throw new Error('Missing email_subject_template');
        }
        try {
            // If the subject template contains templating syntax and subject_template_value exists,
            // format the subject with the subject_template_value
            if (this.row.email_subject_template.includes('<?=') && this.row.subject_template_value) {
                const tpl = HtmlService.createTemplate(this.row.email_subject_template);
                // Add subject_template_value directly to the template
                tpl.subject_template_value = this.row.subject_template_value;
                // Add other template values
                Object.assign(tpl, this.getTemplateValues());
                let s = tpl.evaluate().getContent().trim();
                return TextUtils.decodeHtmlEntities(s);
            }
            else {
                // If no templating is needed or no subject_template_value provided,
                // just return the subject template as is
                return this.row.email_subject_template;
            }
        }
        catch (e) {
            AppScriptLogger.error('Error building subject', e);
            throw e; // Re-throw the error to prevent the email from being sent
        }
    }
    getTemplateValues() {
        const json = TextUtils.sanitizeJsonText(this.row.email_template_values);
        const vals = JSON.parse(json);
        const sid = EmailUtils.parseSessionId(this.row.revu_session_invite);
        if (sid)
            vals.sessionId = sid;
        return vals;
    }
    getAttachments() {
        if (!this.row.attachments_urls)
            return [];
        const blobs = [];
        for (const url of this.row.attachments_urls.split(/[,;]+/).map(s => s.trim())) {
            try {
                const id = FileUtils.extractFileId(url);
                if (FileUtils.isFileTooLarge(id, MAX_ATTACHMENT_SIZE)) {
                    throw new Error(`Attachment too large: ${id}`);
                }
                blobs.push(FileUtils.getFileBlob(id));
            }
            catch (e) {
                AppScriptLogger.error(`Error attaching ${url}`, e);
                throw e;
            }
        }
        return blobs;
    }
    trashAttachments() {
        if (!this.row.attachments_urls)
            return;
        for (const url of this.row.attachments_urls.split(/[,;]+/).map(s => s.trim())) {
            try {
                FileUtils.trashFile(FileUtils.extractFileId(url));
            }
            catch { }
        }
    }
}
/* ------------------------------------------------------------------ */
/*                       TEMPLATE  MANAGEMENT                         */
/* ------------------------------------------------------------------ */
class TemplateManager {
    constructor(ss) {
        this.ss = ss;
        this.templateData = [];
        this.headerMap = {};
        this.index = {};
        this.templatesSheet = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TEMPLATES);
        this.loadTemplateData();
    }
    loadTemplateData() {
        const rows = this.templatesSheet.getDataRange().getValues();
        if (rows.length < 2)
            return;
        this.headerMap = SpreadsheetUtils.mapHeadersToIndices(rows[0]);
        for (let i = 1; i < rows.length; i++) {
            const label = rows[i][this.headerMap['distribution_template_label']];
            if (label) {
                this.index[label] = i;
                this.templateData[i] = rows[i];
            }
        }
    }
    getTemplateByLabel(label) {
        const idx = this.index[label];
        if (idx == null)
            return null;
        return SpreadsheetUtils.mapRowToObject(this.templateData[idx], this.headerMap, Object.keys(this.headerMap));
    }
    getAvailableTemplates() {
        return Object.keys(this.index);
    }
}
/* ------------------------------------------------------------------ */
/*                         EMAIL  PROCESSOR                           */
/* ------------------------------------------------------------------ */
/** EmailProcessor handles the main logic for sending emails and applying templates */
class EmailProcessor {
    constructor() {
        this.ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        this.source = this.ss.getSheetByName(SheetNames.TO_SEND);
        this.history = SpreadsheetUtils.getOrCreateSheet(this.ss, SheetNames.SENT_HISTORY, [...this.getHeaderNames(), 'datetime']);
        this.tm = new TemplateManager(this.ss);
    }
    getHeaderNames() {
        return this.source.getRange(1, 1, 1, this.source.getLastColumn()).getValues()[0].filter(Boolean);
    }
    getHeaderMap() {
        return SpreadsheetUtils.mapHeadersToIndices(this.source.getRange(1, 1, 1, this.source.getLastColumn()).getValues()[0]);
    }
    /** sendEmails processes and moves rows to history */
    sendEmails() {
        const data = this.source.getDataRange().getValues();
        const hMap = this.getHeaderMap();
        for (let i = data.length - 1; i >= 1; i--) {
            const row = data[i];
            if (row.every(c => !c))
                continue;
            let emailRow;
            try {
                emailRow = SpreadsheetUtils.mapRowToObject(row, hMap);
                // apply template
                if (emailRow.distribution_template_label) {
                    const tpl = this.tm.getTemplateByLabel(emailRow.distribution_template_label);
                    if (!tpl)
                        throw new Error(`Template "${emailRow.distribution_template_label}" not found`);
                    for (const k in tpl) {
                        if (!emailRow[k] && tpl[k])
                            emailRow[k] = tpl[k];
                    }
                }
            }
            catch (e) {
                AppScriptLogger.error(`Template error row ${i + 1}`, e);
                SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference(`Error with template in row ${i + 1}`, e));
                continue;
            }
            try {
                const builder = new EmailBuilder(emailRow);
                if (builder.sendEmail()) {
                    this.history.appendRow([...row, new Date()]);
                    this.source.deleteRow(i + 1);
                }
            }
            catch (e) {
                AppScriptLogger.error(`Error processing row ${i + 1}`, e);
            }
        }
    }
    /** applyTemplatesToPendingRows collects errors and updates empty cells */
    applyTemplatesToPendingRows() {
        const data = this.source.getDataRange().getValues();
        const hMap = this.getHeaderMap();
        let updated = 0;
        const errors = [];
        for (let i = 1; i < data.length; i++) {
            const sheetRow = data[i];
            const rowIndex = i + 1;
            const er = SpreadsheetUtils.mapRowToObject(sheetRow, hMap);
            if (!er.distribution_template_label)
                continue;
            const tpl = this.tm.getTemplateByLabel(er.distribution_template_label);
            if (!tpl) {
                errors.push(`Row ${rowIndex}: template "${er.distribution_template_label}" not found`);
                continue;
            }
            let changed = false;
            for (const k in tpl) {
                const col = hMap[k];
                if (col == null)
                    continue;
                const val = sheetRow[col];
                if ((!val || val === '') && tpl[k]) {
                    this.source.getRange(rowIndex, col + 1).setValue(tpl[k]);
                    changed = true;
                }
            }
            if (changed)
                updated++;
        }
        return { updatedRowCount: updated, errors };
    }
}
/* ------------------------------------------------------------------ */
/*                         GLOBAL  FUNCTIONS                          */
/* ------------------------------------------------------------------ */
/** Applies templates to pending rows and shows a consolidated alert */
function applyTemplatesToPendingRows() {
    try {
        AppScriptLogger.info('Applying templates to pending rows...');
        const res = new EmailProcessor().applyTemplatesToPendingRows();
        AppScriptLogger.info(`Template application completed. Updated ${res.updatedRowCount} rows.`);
        const ui = SpreadsheetApp.getUi();
        if (res.errors.length) {
            const msg = AppScriptLogger.formatErrorWithExecutionLogReference(`Some errors occurred while applying templates:\n\n${res.errors.join('\n')}`, {});
            ui.alert(msg);
        }
    }
    catch (e) {
        AppScriptLogger.error('Error applying templates', e);
        SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error applying templates', e));
    }
}
/** Processes all pending email distributions */
function processEmailDistributions() {
    try {
        AppScriptLogger.info('Processing email distributions...');
        new EmailProcessor().sendEmails();
        AppScriptLogger.info('Email distribution completed.');
    }
    catch (e) {
        AppScriptLogger.error('Error processing email distributions', e);
        SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error processing email distributions', e));
    }
}
/** Initializes spreadsheet structure */
function initializeSpreadsheetStructure() {
    try {
        AppScriptLogger.info('Initializing spreadsheet structure...');
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const std = [
            'distribution_template_label', 'distribution_emails', 'additional_emails',
            'revu_session_invite', 'email_template_values', 'email_body_template',
            'attachments_urls', 'email_subject_template', 'subject_template_value'
        ];
        const toSend = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TO_SEND, std);
        const hist = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.SENT_HISTORY, [...std, 'datetime']);
        const tmpl = SpreadsheetUtils.getOrCreateSheet(ss, SheetNames.TEMPLATES, std);
        [toSend, hist, tmpl].forEach(sh => {
            if (sh.getLastRow() > 0) {
                sh.getRange(1, 1, 1, sh.getLastColumn())
                    .setBackground('#f3f3f3')
                    .setFontWeight('bold');
            }
        });
        SpreadsheetApp.getUi().alert('Spreadsheet structure initialized successfully.');
    }
    catch (e) {
        AppScriptLogger.error('Error initializing spreadsheet structure', e);
        SpreadsheetApp.getUi().alert(AppScriptLogger.formatErrorWithExecutionLogReference('Error initializing spreadsheet structure', e));
    }
}
/** Builds the custom menu on open */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Email Distributions')
        .addItem('Send Pending Emails', 'processEmailDistributions')
        .addItem('Apply Templates to Pending Rows', 'applyTemplatesToPendingRows')
        .addItem('Initialize Spreadsheet Structure', 'initializeSpreadsheetStructure')
        .addToUi();
}
