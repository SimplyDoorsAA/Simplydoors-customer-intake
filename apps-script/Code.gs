/**
 * SimplyDoors project intake form backend.
 *
 * Deployment: paste this into the Apps Script project bound to the intake
 * spreadsheet (Extensions > Apps Script), then Deploy > Manage deployments >
 * edit the existing web app deployment > New version, so the live /exec URL
 * served to index.html actually picks up these changes. Editing the code
 * alone does not update a deployment already in use.
 */

const APP_TOKEN = 'SimplyDoors2026-Secure';
const NOTIFY_EMAIL = 'adem@simplydoors.com';
// Sheet tab submissions are written to. If no sheet with this name exists,
// falls back to whichever sheet is currently active (see getSubmissionsSheet)
// so this change can't orphan existing data in a differently-named tab.
const SHEET_NAME = 'Submissions';

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
      companyName: sanitizeCell(e.parameter.companyName),
      contactName: sanitizeCell(e.parameter.contactName),
      phone: sanitizeCell(e.parameter.phone),
      email: sanitizeCell(e.parameter.email),
      billingAddress: sanitizeCell(e.parameter.billingAddress),
      jobSiteAddress: sanitizeCell(e.parameter.jobSiteAddress),
      projectScope: sanitizeCell(e.parameter.projectScope)
    };

    // Saving the submission and notifying the team are separate concerns: a
    // customer's data being recorded must never depend on whether the
    // notification email happens to succeed (mail quota, a transient API
    // error, a revoked authorization). Save first, notify second, and never
    // let a mail failure make a successfully-saved submission look failed.
    const rowValues = [
      new Date(),
      data.companyName,
      data.contactName,
      data.phone,
      data.email,
      data.billingAddress,
      data.jobSiteAddress,
      data.projectScope,
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

    const emailSent = notifyTeam(data);
    sheet.getRange(rowIndex, statusColumn).setValue(emailSent ? 'sent' : 'failed');

    return jsonResponse({ result: 'success' });
  } catch (error) {
    return jsonResponse({ result: 'error', message: error.message });
  }
}

function notifyTeam(data) {
  const subject = `New Project Intake: ${data.companyName || 'Unknown Builder'}`;
  const body = [
    'You have received a new project intake submission.',
    '',
    'COMPANY DETAILS',
    `Builder: ${data.companyName || 'N/A'}`,
    `Contact: ${data.contactName || 'N/A'}`,
    `Phone: ${data.phone || 'N/A'}`,
    `Email: ${data.email || 'N/A'}`,
    '',
    'PROJECT DETAILS',
    `Billing Address: ${data.billingAddress || 'N/A'}`,
    `Job Site Address: ${data.jobSiteAddress || 'N/A'}`,
    '',
    'SCOPE OF WORK:',
    data.projectScope || 'N/A'
  ].join('\n');

  const options = { to: NOTIFY_EMAIL, subject: subject, body: body };
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;

  // No sheet named SHEET_NAME exists yet. Rather than risk creating a
  // second, empty tab alongside your existing data, keep using whichever
  // sheet is active — the same behavior as before this change. Rename your
  // submissions tab to "Submissions" (or update SHEET_NAME) so this
  // resolves reliably even if a different tab was last viewed in the UI.
  return spreadsheet.getActiveSheet();
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
