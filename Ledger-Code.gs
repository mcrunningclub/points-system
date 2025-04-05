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

/**
 * Removed functions from previous static version (Points Ledger 2023/2024)
 *  - appendHeadRun() 
 *  - importMembers()
 *  - tallyPointsOnce()
 *  - updateHeadRunPoints()
 *  - formatDateNote()
 * 
 * UPDATED: MARCH 27, 2025
 */


// CURRENTLY IN REVIEW!
function newSubmission() {
  formatSpecificColumns();
  sortTimestampByAscending();
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

function getLatestSubmissionTimestamp() {
  return getSubmissionTimestamp(getValidLastRow(LOG_SHEET));
}

function getSubmissionTimestamp(row) {
  const sheet = LOG_SHEET;
  const timestampCol = LOG_INDEX.EVENT_TIMESTAMP;
  const timestamp = sheet.getRange(row, timestampCol).getValue();
  return new Date(timestamp);
}


/**
 * Find row index of last entry, starting from bottom using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Target sheet.
 * @return {integer}  Returns 1-index of last row in `sheet`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sept 1, 2024
 * @update  Mar 23, 2025
 */

function getValidLastRow(sheet) {
  let lastRow = sheet.getLastRow();

  while (sheet.getRange(lastRow, 1).getValue() == "") {
    lastRow = lastRow - 1;
  }
  return lastRow;
}


// Return latest log values
function getLatestLog_() {
  return getLogInRow_();
}

function getLogInRow_(row = getValidLastRow(LOG_SHEET)) {
  const sheet = LOG_SHEET;
  const numCols = sheet.getLastColumn();
  return sheet.getSheetValues(row, 1, 1, numCols)[0];
}

function getAttendeesInLog_(row) {
  // Get log attendees using stored index
  const attendeesCol = LOG_INDEX.ATTENDEE_NAME_EMAIL - 1;
  const thisLog = getLogInRow_(row);

  // Return log attendees
  return thisLog[attendeesCol];
}

function getMapUrlInRow_(row) {
  return getLogCell_(row, LOG_INDEX.MAP_URL) || "";
}

function getEventPointsInRow_(row) {
  return getLogCell_(row, LOG_INDEX.EVENT_POINTS) || 0;
}

function getLogCell_(row, column) {
  const sheet = LOG_SHEET;
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
  const pointSheet = LEDGER_SHEET;

  // Define dimensions of sheet data
  const startCol = 1;
  const startRow = 2;
  const numRows = getValidLastRow(pointSheet) - 1;   // Remove header row

  return pointSheet.getSheetValues(startRow, startCol, numRows, numCols);
}


function getLedgerEntry(email, ledgerData) {
  const row = findMemberInLedger_(email, ledgerData);
  return ledgerData[row] ?? [];
}

/**
 * Recursive function to search for entry by email in `sheet` using binary search.
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} emailToFind  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * @param {number} [start=2]  The starting row index for the search (1-indexed). 
 *                            Defaults to 2 (the second row) to avoid the header row.
 * @param {number} [end=MASTER_SHEET.getLastRow()]  The ending row index for the search. 
 *                                                  Defaults to the last row in the sheet.
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

function findMemberInLedger_(emailToFind, ledger) {
  const EMAIL_COL = LEDGER_INDEX.EMAIL - 1;   // Make 0-indexed
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
    if (emailAtMid === emailToFind) {
      return mid;

    /** If the email at the middle row is alphabetically smaller, search the right half. */
    /** Note: use localeString() to ensure string comparison matches GSheet. */
    } else if (emailAtMid.localeCompare(emailToFind) === -1) {
      return findThisEmailBinarySearch(mid + 1, end);

    /** If the email at the middle row is alphabetically larger, search the left half. */
    } else {
      return findThisEmailBinarySearch(start, mid - 1);
    }
  };
}

