/**
 * RSVP receiver for the Krishanjeet & Aanchal website.
 * Appends every form submission as a new row in the Google Sheet.
 *
 * Setup (one time):
 *  1. Open the sheet → Extensions → Apps Script.
 *  2. Replace the editor contents with this file and click Save.
 *  3. Deploy → New deployment → type "Web app".
 *       Execute as: Me      Who has access: Anyone
 *  4. Authorize when asked, then copy the Web app URL (ends in /exec)
 *     into RSVP_SHEET_URL in index.html.
 *  If you edit this script later: Deploy → Manage deployments → Edit → Version: New version.
 */

const SHEET_ID = '1s7WRxHR-twCDT_xCYt4qx3UiANc3IgJU8JieYOssnuk';
const HEADERS = ['Timestamp', 'Name', 'Mobile Number', 'Attending Reception', 'Who Will Get Emotional First', 'Note'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    // Leading apostrophe keeps phone numbers as text (no lost leading zeros / + sign).
    sheet.appendRow([
      new Date(),
      asText(data.guestName),
      data.mobileNumber ? "'" + data.mobileNumber : '',
      asText(data.attendingReception),
      asText(data.guessEmotionalFirst),
      asText(data.message)
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Stops guest input like "=..." from being run as a spreadsheet formula.
function asText(value) {
  const s = String(value || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
