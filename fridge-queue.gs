// These strings must match the column headers in the first row of the spreadsheet.
const SAMPLE_ID = "sample ID"
const EMAIL = "email"
const QUEUED = "queued"
const UNLOAD_SCHEDULED = "unload scheduled"

function installableOnEdit(event) {
  /** Installable trigger configured to fire when any cell or range is edited.
   *
   * This CANNOT be named onEdit(), or else it will fire twice: once as a simple trigger,
   * and again as an installable trigger.
   *
   * This script should only autofill values once, when the sampleID is typed or pasted.
   * This is achieved by never overwriting values: once autofilled, always filled.
   *
   * Note: we cannot rely on event.value and event.oldValue to distinguish
   * inserts, changes, and deletions, because:
   *   1. Both are null for multi-cell edits.
   *   2. Both are null for copy-paste edits.
   */
  // do nothing on deletions and multi-column edits
  if (event.range.isBlank() || event.range.getNumColumns() != 1) {
    return;
  }
  
  if (event.range.getColumn() == getColumnByName(event.range, SAMPLE_ID)) {
    autofill(event);
  } else if (event.range.getColumn() == getColumnByName(event.range, UNLOAD_SCHEDULED)) {
    notify(event);
  }
}

function autofill(event) {
  /** Autofill fields EMAIL and QUEUED. */
  var email_col = getColumnByName(event.range, EMAIL);
  var queued_col = getColumnByName(event.range, QUEUED);
  
  var row = event.range.getRow();
  var numRows = event.range.getNumRows();
  
  var email_range = event.range.getSheet().getRange(row, email_col, numRows);
  var queued_range = event.range.getSheet().getRange(row, queued_col, numRows);
  
  safelySetValue(email_range, event.user.getEmail());
  safelySetValue(queued_range, new Date());
}

function notify(event) {
  /** Notify the next person in the queue that unload has been scheduled.
   *
   * Increment until an email is found without a corresponding UNLOAD_SCHEDULED date.
   */
  var sampleIdCol = getColumnByName(event.range, SAMPLE_ID);
  var emailCol = getColumnByName(event.range, EMAIL);
  var unloadScheduledCol = getColumnByName(event.range, UNLOAD_SCHEDULED);
  
  var sheet = event.range.getSheet();
  var currentRow = event.range.getRow();
  var currentUser = sheet.getRange(currentRow, emailCol).getValue();
  var unloadTime = sheet.getRange(currentRow, unloadScheduledCol).getValue();
  
  var lastRow = sheet.getLastRow();
  for (row = currentRow + 1; row <= lastRow; row++) {
    var nextUser = sheet.getRange(row, emailCol).getValue();
    var nextUnloadTime = sheet.getRange(row, unloadScheduledCol).getValue();
    var nextSample = sheet.getRange(row, sampleIdCol).getValue();
    var sheetName = sheet.getName();
    if (nextUser != "" && nextUnloadTime == "") {
      subject = `[fridge-queue/${sheetName}] ${currentUser.split("@")[0]} `
        + `is scheduled to unload ${Utilities.formatDate(new Date(unloadTime), 'America/New_York', 'M/d h:mma')}`;
      
      var options = {
        noReply: true,
        htmlBody: "You're next in the queue."
        + `<blockquote><b>fridge:</b> ${sheetName}<br/>`
        + `<b>sample:</b> ${nextSample}<br/>`
        + `<b>scheduled:</b> ${unloadTime}</blockquote>`
        + `Coordinate the exact time with ${currentUser}.`
        + `See <a href="${sheet.getParent().getUrl()}">queue here</a>.<br/><br/>`
        + `This is an automated message.`,
      };
      MailApp.sendEmail(nextUser, subject, '', options)
      break;
    }
  }
}

function highlight(event) {
  /** Highlight the currently loaded sample. */
  return;
}

function getColumnByName(range, colName) {
  /** Get the number of the column with "colName" in the first row.
   *
   * Note: columns are 1-indexed but javascript is 0-indexed.
   *
   * TODO Refactor: this queries the spreadsheet every time and is expensive.
   */
  return range.getSheet().getRange("1:1").getValues()[0].indexOf(colName) + 1;
}

function safelySetValue(range, newVal) {
  /** Set value only if it will not overwrite anything. */
  if (range.isBlank()) {
    range.setValue(newVal);
  }
}
