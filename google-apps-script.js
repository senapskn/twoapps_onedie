// Roll and Tell – receives answers from form.html and writes them to this Google Sheet.
// Paste this into Extensions → Apps Script, then Deploy → New deployment → Web app
// (Execute as: Me, Who has access: Anyone). Copy the /exec URL into SUBMIT_URL in form.html.

const SHEET_NAME = 'Responses';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(data.header);
      sh.setFrozenRows(1);
    }

    // Ignore a submission that is already in the sheet (the form may send twice).
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat();
      if (ids.indexOf(data.id) !== -1) return reply({ ok: true, duplicate: true });
    }

    if (data.rows && data.rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, data.rows.length, data.rows[0].length).setValues(data.rows);
    }
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply({ ok: true, message: 'Roll and Tell receiver is running.' });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
