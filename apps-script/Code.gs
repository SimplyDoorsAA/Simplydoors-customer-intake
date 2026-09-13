/**
 * SimplyDoors project intake form backend.
 *
 * Writes to the "SimplyDoors Project Intake" spreadsheet by ID (see
 * SPREADSHEET_ID below), so this works whether the Apps Script project is
 * bound to that sheet or standalone.
 *
 * Deployment: paste this into whichever Apps Script project /exec URL is
 * referenced by SCRIPT_URL in index.html, then Deploy > Manage deployments >
 * edit the existing web app deployment > New version, so that URL actually
 * picks up these changes. Editing the code alone does not update a
 * deployment already in use.
 */

const APP_TOKEN = 'SimplyDoors2026-Secure';
const NOTIFY_EMAIL = 'adem@simplydoors.com';
// The "SimplyDoors Project Intake" spreadsheet. Addressed explicitly by ID
// rather than via SpreadsheetApp.getActiveSpreadsheet() — a standalone
// script project (one opened from script.google.com or Drive rather than
// via Extensions > Apps Script from inside the sheet) has no "active
// spreadsheet" at all and getActiveSpreadsheet() returns null, so this
// works regardless of whether this project is bound to the sheet or not.
const SPREADSHEET_ID = '19wstlrr2cPf3_hvW836lOFqQ_faXvV_jq9jBx0ySBzo';
// Sheet tab submissions are written to. If no sheet with this name exists,
// falls back to the spreadsheet's first sheet (see getSubmissionsSheet).
const SHEET_NAME = 'Submissions';
// MailApp/Gmail caps a single outgoing message (body + attachments combined)
// at roughly 25MB. Stay comfortably under that so a large-but-under-client-
// limit upload still can't produce an email Gmail silently refuses to send.
const MAX_ATTACHMENTS_BYTES = 20 * 1024 * 1024;

function doPost(e) {
  try {
    if (!e || !e.parameter || e.parameter.appToken !== APP_TOKEN) {
      return jsonResponse({ result: 'error', message: 'Unauthorized request origin.' });
    }

    // Honeypot: a filled hidden field means a bot. Return a fake success so
    // it doesn't learn it was caught, but record and email nothing.
    if (e.parameter.botTrap) {
      return jsonResponse({ result: 'success' });
    }

    const data = {
      contactName: sanitizeCell(e.parameter.contactName),
      companyName: sanitizeCell(e.parameter.companyName),
      phone: sanitizeCell(e.parameter.phone),
      email: sanitizeCell(e.parameter.email),
      projectAddress: sanitizeCell(e.parameter.projectAddress),
      projectDescription: sanitizeCell(e.parameter.projectDescription)
    };

    const attachments = collectAttachments(e);
    const attachmentNames = attachments.map(function (file) { return file.getName(); });

    // Saving the submission and notifying the team are separate concerns: a
    // customer's data being recorded must never depend on whether the
    // notification email happens to succeed (mail quota, a transient API
    // error, a revoked authorization). Save first, notify second, and never
    // let a mail failure make a successfully-saved submission look failed.
    const rowValues = [
      new Date(),
      data.contactName,
      data.companyName,
      data.phone,
      data.email,
      data.projectAddress,
      data.projectDescription,
      sanitizeCell(attachmentNames.join(', ')),
      'pending' // email status, overwritten below once we know the outcome
    ];
    const statusColumn = rowValues.length;

    const sheet = getSubmissionsSheet();
    let rowIndex;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.appendRow(rowValues);
      rowIndex = sheet.getLastRow();
    } finally {
      lock.releaseLock();
    }

    const emailSent = notifyTeam(data, attachments);
    sheet.getRange(rowIndex, statusColumn).setValue(emailSent ? 'sent' : 'failed');

    return jsonResponse({ result: 'success' });
  } catch (error) {
    return jsonResponse({ result: 'error', message: error.message });
  }
}

// The client sends each selected file under its own field name
// (attachment_0, attachment_1, ...) rather than repeating one field name,
// since e.files does not reliably expose more than one blob per field name.
function collectAttachments(e) {
  if (!e.files) return [];

  const attachments = [];
  let totalBytes = 0;

  Object.keys(e.files)
    .filter(function (name) { return name.indexOf('attachment_') === 0; })
    .sort() // attachment_0, attachment_1, ... — keep the user's selection order
    .forEach(function (name) {
      const blob = e.files[name];
      totalBytes += blob.getBytes().length;
      if (totalBytes > MAX_ATTACHMENTS_BYTES) {
        throw new Error('Attachments are too large to email. Please resend with less than 20MB total.');
      }
      attachments.push(blob);
    });

  return attachments;
}

function notifyTeam(data, attachments) {
  const subject = `New Project Intake: ${data.contactName || 'Unknown Contact'}`;
  const bodyLines = [
    'You have received a new project intake submission.',
    '',
    'CONTACT DETAILS',
    `Name: ${data.contactName || 'N/A'}`,
    `Company: ${data.companyName || 'N/A'}`,
    `Phone: ${data.phone || 'N/A'}`,
    `Email: ${data.email || 'N/A'}`,
    '',
    'PROJECT DETAILS',
    `Project Address: ${data.projectAddress || 'N/A'}`,
    '',
    'PROJECT DESCRIPTION:',
    data.projectDescription || 'N/A'
  ];

  if (attachments.length > 0) {
    bodyLines.push('', `ATTACHMENTS (${attachments.length}): see attached files.`);
  }

  const options = { to: NOTIFY_EMAIL, subject: subject, body: bodyLines.join('\n') };
  if (attachments.length > 0) {
    options.attachments = attachments;
  }
  if (isValidEmail(data.email)) {
    options.replyTo = data.email;
  }

  try {
    MailApp.sendEmail(options);
    return true;
  } catch (mailError) {
    // Quota exceeded or a transient failure shouldn't take the whole request
    // down with it — the submission above is already saved. Logged here so
    // it's visible in Executions, and marked in the sheet's status column.
    Logger.log('Notification email failed: ' + mailError.message);
    return false;
  }
}

function getSubmissionsSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;

  // No sheet named SHEET_NAME exists yet. Rather than risk creating a
  // second, empty tab alongside your existing data, use the spreadsheet's
  // first sheet. Rename your submissions tab to "Submissions" (or update
  // SHEET_NAME) so this resolves by name instead of position.
  return spreadsheet.getSheets()[0];
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value || '');
}

function sanitizeCell(value) {
  const text = (value || '').toString();
  // Google Sheets treats a leading =, +, -, or @ as a formula regardless of
  // how the cell was written. Prefix with an apostrophe to force plain text
  // so a submitted field value can't inject a spreadsheet formula.
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
