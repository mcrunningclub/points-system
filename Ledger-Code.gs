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
  const sheet = LOG_SHEET;
  const timestampCol = LOG_INDEX.EVENT_TIMESTAMP;
  const lastRow = getValidLastRow(sheet);

  const timestamp = sheet.getRange(lastRow, timestampCol).getValue();
  return new Date(timestamp);
}


/**
 * Find row index of last entry, starting from bottom using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in `sheet`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sept 1, 2024
 * @update  Mar 23, 2025
 */

function getValidLastRow(sheet) {
  let lastRow = sheet.getLastRow();
  
  while (sheet.getRange(lastRow, 0).getValue() == "") {
    lastRow = lastRow - 1;
  }
  return lastRow;
}


function getLedgerEntry(email, ledgerData) {
  const row = findMemberInLedger(email, ledgerData);
  return ledgerData[row];
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
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Mar 23, 2025
 * @update  Mar 23, 2025
 * 
 * @example `const submissionRowNumber = findMemberByBinarySearch('example@mail.com', getLedgerData());`
 */

function findMemberInLedger(emailToFind, ledger) {
  const EMAIL_COL = LEDGER_INDEX.EMAIL - 1;   // Make 0-indexed
  return findThisEmailBinarySearch();

  /** Define as inner-function to prevent passing `emailToFind` and `ledger` at every call */ 
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


// OLD FUNCTIONS

function newSubmission() {
  formatSpecificColumns();
  sortNameByAscending();
  updateHeadRunPoints();
}



/**
 * Gets the names of the attendees by head run and updates their points.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 29, 2023
 * @update  Nov 13, 2023
 */

function updateHeadRunPoints() {
  // `Head Run Attendance` Google Sheet
  var ss = LEDGER_SS;

  const TIMESTAMP_COL = 1;
  const HEAD_RUN_COL = 4;
  const ATTENDEES_COL = 5;
  const IS_ADDED_COL = 6;
  const NAMES_NOT_FOUND_COL = 7;

  const sheetImport = ss.getSheetByName("Head Run Attendance");
  const timestamp = sheetImport.getRange(sheetImport.getLastRow(), TIMESTAMP_COL).getValue();
  const headRun = sheetImport.getRange(sheetImport.getLastRow(), HEAD_RUN_COL).getValue();
  const isAddedRange = sheetImport.getRange(sheetImport.getLastRow(), IS_ADDED_COL);
  const unfoundNameRange = sheetImport.getRange(sheetImport.getLastRow(), NAMES_NOT_FOUND_COL);

  if (isAddedRange.getValue()) return;    // Exit since data has already been added to points ledger

  const note = formatDateNote(timestamp, headRun);
  Logger.log(note);

  // Get Attendees
  const rangeAttendees = sheetImport.getRange(sheetImport.getLastRow(), ATTENDEES_COL);
  var attendees = rangeAttendees.getValue();

  var attendeesArray = attendees.split('\n');  // split the string into an array;
  attendeesArray = attendeesArray.map(str => str.trim());   // trim whitespace from every string in array
  
  var cannotFindArray = [];

  for(var i=0; i<attendeesArray.length; i++) {
    var name = attendeesArray[i];
    if (name.toLowerCase() === "none") break;

    var error = appendHeadRun(name, note);
    if (error === -2) cannotFindArray.push(name);
  }

  // Add names not found from attendance sheet
  const cellValue = cannotFindArray.join(", ");
  unfoundNameRange.setValue(cellValue);
  
  isAddedRange.check();
}


/**
 * Tally all the points of head run attendees from the beginning.
 * 
 * @warning ONLY USE TO IMPORT POINTS TO A BLANK LEDGER SHEET.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 25, 2023
 * @update  Nov 25, 2023
 */

function tallyPointsOnce() {
  var sheet = LEDGER_SS;

  // Constants in `Head Run Attendance` Sheet
  const TIMESTAMP_COL = 1;
  const HEAD_RUN_COL = 4;
  const ATTENDEES_COL = 5;
  const NAMES_NOT_FOUND_COL = 7;

  // Create array with all referencs to attendance submissions
  const attendanceSheet = sheet.getSheetByName("Head Run Attendance");
  const rowNum = attendanceSheet.getLastRow();
  const attendances = attendanceSheet.getRange(2, 1, rowNum, ATTENDEES_COL).getValues();   // attendance array

  for(var i=0; i< attendances.length; i++) {
    var row = attendances[i];
    var timestamp = row[TIMESTAMP_COL -1];
    var headRun = row[HEAD_RUN_COL -1];

    // Create note to append
    const note = formatDateNote(timestamp, headRun);
    Logger.log(note);

    // Get Attendees
    var attendees = row[ATTENDEES_COL -1];
    var attendeesArray = attendees.split('\n');  // split the string into an array;
    attendeesArray = attendeesArray.map(str => str.trim());   // trim whitespace from every string in array

    var cannotFindArray = [];

    for(var j=0; j<attendeesArray.length; j++) {
      var name = attendeesArray[j];
      if (name.toLowerCase() === "none") break;

      var error = appendHeadRun(name, note);
      if (error === -2) cannotFindArray.push(name);
    }

    // Add names not found from attendance sheet
    const cellValue = cannotFindArray.join(", ");

    const unfoundNamesCell = attendanceSheet.getRange(i + 2, NAMES_NOT_FOUND_COL);   // reference to cell    
    unfoundNamesCell.setValue(cellValue);
  }
}


/**
 * Imports members from `Membership Collected` to `Points Ledger`
 * @warning ONLY USE ONCE!
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 20, 2023
 * @update  Nov 20, 2023
 */

function importMembers() {
  const MEMBERSHIP_SPREADSHEET = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/edit?usp=sharing");
  
  const copySheet = MEMBERSHIP_SPREADSHEET.getSheetByName('Fall 2023');   // Fall 2023 member list
  const NAME_COL = 3;
  const FEE_PAID_COL = 14;
  const MEMBER_ID_COL = 20;

  const pasteSheet = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);
  const PASTE_MEMBER_ID_COL = 1;
  const PASTE_FEE_PAID_COL = 2;
  const PASTE_FULL_NAME_COL = 3;

  // Get size of source `Fall 2023` list
  const numCols = copySheet.getLastColumn();
  const numRows = copySheet.getLastRow();

  // Set member id in `Member Points`
  const copyMemberIDRange = copySheet.getRange(2, MEMBER_ID_COL, numRows, 1);
  const pasteMemberIDRange = pasteSheet.getRange(2, PASTE_MEMBER_ID_COL, numRows, 1);
  pasteMemberIDRange.setValues(copyMemberIDRange.getValues());

  // Set fee paid in `Member Points`
  const copyFeePaidRange = copySheet.getRange(2, FEE_PAID_COL, numRows, 1);
  const pasteFeePaidRange = pasteSheet.getRange(2, PASTE_FEE_PAID_COL, numRows, 1);
  pasteFeePaidRange.setValues(copyFeePaidRange.getValues());

  // Set full name in `Member Points`
  const rangeName = copySheet.getRange(2, NAME_COL, numRows, 2);
  var members = rangeName.getValues();

  for (var rowNum=0; rowNum < members.length; rowNum++) {
    // Get first and last names of member at `rowNum`
    var firstName = members[rowNum][0];
    var lastName = members[rowNum][1];

    // Get full name and copy to pasteSheet
    var fullName = firstName + " " + lastName;
    pasteSheet.getRange(rowNum + 2, PASTE_FULL_NAME_COL).setValue(fullName);
  }
  return;

  // DEAD CODE
  // Get `SpreadsheetApp.Range` object
  const sourceRow = copySheet.getRange(lastRow, 1, 1, numCols);
  var rangeBackup = pasteSheet.getRange(pasteSheet.getLastRow() +1, 1, 1, numCols);
  
  // Copy and paste values
  const valuesToCopy = sourceRow.getValues();
  rangeBackup.setValues(valuesToCopy);

  return;
}


/**
 * Appends head run event using name as key and note in `Member Points`.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2023
 * @update  Oct 30, 2023
 */

function appendHeadRun(nameKey, note) {
  const sheetPoints = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);
  const headRunPoints = 50;

  var data = sheetPoints.getDataRange().getValues();
  var rowIndex = -1;
  var lastColumnIndex = -1;

  // Find the row with the specified key
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === nameKey) {   // Check `Full Name` (column C)
      rowIndex = i;
      break;
    }
  }

  // TODO: Save Index of Last Used Column

  if (rowIndex !== -1) {
    // Find last non-empty column
    for (var j = data[0].length - 1; j >= 0; j--) {
      if (data[rowIndex][j] !== "") {
        lastColumnIndex = j;
        break;
      }
    }

    // Append head run event to ledger
    var targetCell = sheetPoints.getRange(rowIndex + 1, lastColumnIndex + 2);
    targetCell.setValue(headRunPoints);
    targetCell.setNote(note);
    targetCell.setBackground('#f9cb9c');
  }

  else {
    return -2;
  }
}


function test_() {
  //var date = "2023-10-18 20:04:53";
  //var headrun = "Wednesday 6pm (Intermediate Run)";
  
  //Logger.log(formatDateNote(date, headrun));

  //const attendees = "James lee , darius  Corb, carly cake";
  const attendees = "James lee \n darius Corb\ncarly cake ";

  var splitArray = attendees.split('\n');  // split the string into an array;
  if (splitArray.length < 2) splitArray = attendees.split(',');  // split the string into an array;

  // Trim whitespace from strings and set to Title Case
  var formattedAttendees = splitArray.map(str => str.trim());   
  formattedAttendees = formattedAttendees.map(str => toTitleCase(str));
  
  var newValue = formattedAttendees.join('\n');       // combine all array elements into single string
  Logger.log(newValue);
}

