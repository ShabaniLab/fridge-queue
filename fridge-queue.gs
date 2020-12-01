// These strings must match the column headers in the first row of the spreadsheet.
const SAMPLE_ID = "sample ID"
const EMAIL = "email"
const QUEUED = "queued"
const UNLOAD_SCHEDULED = "unload scheduled"

function onEdit(event) {
  /** Triggered when any cell or range is edited.
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
    Logger.log("abort");
    return;
  }
  
  if (event.range.getColumn() == getColumnByName(event.range, SAMPLE_ID)) {
    autofill(event);
  } else if (event.range.getColumn() == getColumnByName(event.range, UNLOAD_SCHEDULED)) {
    notify(event);
  } else {
    Logger.log("nothing to do");
  }
}

function autofill(event) {
  /** Autofill fields EMAIL and QUEUED. */
  Logger.log("autofill");
  var email_col = getColumnByName(event.range, "email");
  var queued_col = getColumnByName(event.range, "queued");
  
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
  Logger.log("notify");
  var lastRow = event.range.getSheet().getLastRow();
  for (row = event.range.getRow() + 1; row <= lastRow; row++) {
    var email = event.range.getSheet().getRange(row, getColumnByName(event.range, EMAIL)).getValue();
    var unloadScheduled = event.range.getSheet().getRange(row, getColumnByName(event.range, UNLOAD_SCHEDULED)).getValue();
    if (email != "" && unloadScheduled == "") {
      Logger.log("sending email to " + email);
      // TODO send email
      break
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
   */
  return range.getSheet().getRange("1:1").getValues()[0].indexOf(colName) + 1;
}

function safelySetValue(range, newVal) {
  /** Set values only if it will not overwrite anything. */
  if (range.isBlank()) {
    range.setValue(newVal);
  } else {
    Logger.log("unsafe to set: not blank");
  }
}
