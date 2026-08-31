/*
Copyright 2024 Jikael Gagnon (for McGill Students Running Club)

Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// CURRENTLY IN REVIEW!
function newSubmission() {
  formatSpecificColumns();
  sortLogsByTimestamp();
  //sendStatsEmail()
}


/**
 * Return latest head run submission timestamp in `LOG_SHEET`.
 * 
 * @return {Date}  Headrun submission timestamp as Date object.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Mar 23, 2025
 */

function getLatestTimestamp_() {
  return getTimestampInRow_(getValidLastRow_(GET_LOG_SHEET()));
}


/**
 * Return timestamp for a specified row in LOG_SHEET.
 * 
 * @param {integer} row  Row number.
 * @return {Date}  Timestamp as Date object.
 */
function getTimestampInRow_(row) {
  const sheet = GET_LOG_SHEET();
  const timestampCol = LOG_COL.EVENT_TIMESTAMP;
  const timestamp = sheet.getRange(row, timestampCol).getValue();
  return new Date(timestamp);
}

/**
 * Return content of latest row in log sheet.
 * 
 * @return {Array}  Values of each column in the last row.
 */
function getLatestLog_() {
  return getLogInRow_(getValidLastRow_(LOG_SHEET));
}

/**
 * Return content of specified row in log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {Array}  Values of each column in the specified row.
 */
function getLogInRow_(row) {
  const sheet = GET_LOG_SHEET();
  const numCols = sheet.getLastColumn();
  return sheet.getSheetValues(row, 1, 1, numCols)[0];
}

/**
 * Return list of attendees in specified row of log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {string}  Attendees, separated by newline.
 */
function getAttendeesInRow_(row) {
  return getLogCell_(row, LOG_COL.ATTENDEES);
}

/**
 * Return date and level (of headruns) in specified row of log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {{string: string}}  Object with keys 'date' and 'level', and corresponding values.
 */
function getDateAndLevelInRow_(row) {
  const rawString = getLogCell_(row, LOG_COL.EVENT);
  return {
    'date' : rawString.match(/^[^\n]*\n(.*)$/i)[1],
    'level' : rawString.match(/(?:Headrun)\s+(\w+)/i)[1],
  }
}

/**
 * Return map URL in specified row of log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {string}  Map URL, or emtpy string if not found.
 */
function getMapUrlInRow_(row) {
  return getLogCell_(row, LOG_COL.MAP_URL) || "";
}

/**
 * Return points for the event in specified row of log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {number}  Number of points, or 0 if not found.
 */
function getEventPointsInRow_(row) {
  return getLogCell_(row, LOG_COL.EVENT_POINTS) || 0;
}

/**
 * Return list of headrunners in specified row of log sheet.
 * 
 * @param {integer} row  Row number.
 * @return {string}  Headrunners, separated by newline.
 */
function getHeadrunnersInRow_(row) {
  return getLogCell_(row, LOG_COL.HEADRUNNERS) || "";
}

/**
 * Returns value of specified cell in the log sheet.
 * 
 * @param {number} row  Row number of cell.
 * @param {number} column  Column number of cell.
 * @return {*}  Cell value.
 */
function getLogCell_(row, column) {
  const sheet = GET_LOG_SHEET();
  return sheet.getRange(row, column).getValue();
}

/**
 * Get ledger data from `LEDGER_SHEET` to send emails.
 * 
 * @param {number} [numCols = LEDGER_COL_COUNT]  The number of rows to get starting from email col. 
 *                                               Defaults to last col before events (`LEDGER_COL_COUNT`).
 * 
 * @return {Object[][]}  Ledger data of col size `numCols`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Mar 23, 2025
 * @update  Mar 23, 2025
 */

function getLedgerData_(numCols = LEDGER_COL_COUNT) {
  const sheet = GET_LEDGER_SHEET();

  // Define dimensions of sheet data
  const startCol = 1;
  const startRow = 2;
  const numRows = getValidLastRow_(sheet) - 1;   // Remove header row

  return sheet.getSheetValues(startRow, startCol, numRows, numCols);
}

/**
 * Get ledger data of member using their email.
 * 
 * @param {string} email  Member email address.
 * @param {Object[][]} ledgerData  Ledger data object, from GET_LEDGER_()
 * @return {Array}  Values of the row corresponding to specified member, or empty array if not found.
 */
function getLedgerEntry_(email, ledgerData) {
  const row = findMemberInLedger_(email, ledgerData);
  return ledgerData[row] ?? [];
}

/**
 * Recursive function to search for entry by email in `sheet` using binary search.
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} email  The email address to search for.
 * @param {Objects[][]} ledger  Array containing rows from ledger sheet.
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @example `const submissionRowNumber = findMemberByBinarySearch('example@mail.com', getLedgerData());`
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Mar 23, 2025
 * @update  Mar 23, 2025
 */

function findMemberInLedger_(email, ledger) {
  const EMAIL_COL = LEDGER_COL.EMAIL - 1;   // Make 0-indexed
  return findThisEmailBinarySearch();

  /** Define as inner function to prevent passing `emailToFind` and `ledger` at every call */
  function findThisEmailBinarySearch(start = 1, end = ledger.length) {
    // Base case: If start index exceeds the end index, the email is not found
    if (start > end) {
      return null;
    }

    // Find the middle point between the start and end indexes
    const mid = Math.floor((start + end) / 2);

    // Get the email value at the middle row
    const emailAtMid = ledger[mid][EMAIL_COL];

    // Compare the target email with the middle email
    /** If the email matches, return the row index in ledger */
    if (emailAtMid === email) {
      return mid;

    /** If the email at the middle row is alphabetically smaller, search the right half. */
    /** Note: use localeString() to ensure string comparison matches GSheet. */
    } else if (emailAtMid.localeCompare(email) === -1) {
      return findThisEmailBinarySearch(mid + 1, end);

    /** If the email at the middle row is alphabetically larger, search the left half. */
    } else {
      return findThisEmailBinarySearch(start, mid - 1);
    }
  };
}


/** 
 * Handles the transfered submission from Attendance Code and adds new row to log sheet.
 * Called from the Attendance Code script.
 * 
 * @param {Array[][]} importArr  Submission array with non-empty run levels.
 * @return {integer}  The newly added row number in Log sheet
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Apr 8, 2025
 * @update  Apr 12, 2025
 */

function storeImportFromAttendanceSheet(importArr) {
  const logSheet = GET_LOG_SHEET();
  const funcName = storeImportFromAttendanceSheet.name;
  logAsPL_('Processing following import...', funcName);
  Logger.log(importArr);

  const row = getValidLastRow_(logSheet) + 1;

  try {
    const importNumRows = importArr.length;
    const importNumCols = importArr[0].length;

    // Print number of rows and columns
    logAsPL_(`Row count: ${importNumRows}\tCol count: ${importNumCols}`, funcName, false);
    
    // Now set import as-if (processing occured in Attendance Sheet)
    logSheet.getRange(row, 1, importNumRows, importNumCols).setValues(importArr);

    // Log success message
    logAsPL_(`Successfully imported values to row ${row} in Log Sheet`, funcName, false);
  }
  catch (e) {
    logAsPL_("Unable to fully process 'importArr'", funcName);
    throw e;
  }

  return row;
}

