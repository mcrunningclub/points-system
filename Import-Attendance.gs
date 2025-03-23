/*
Copyright 2023 Andrey Gonzalez (for McGill Students Running Club)

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


/* SHEET CONSTANTS */
const SHEET_NAME = 'Member Points';
const LEDGER_SHEET = SpreadsheetApp.getActiveSpreadsheet();
const TIMEZONE = getUserTimeZone();

function getUserTimeZone() {
  return Session.getScriptTimeZone();
}

function newSubmission() {
  formatSpecificColumns();
  sortNameByAscending();
  updateHeadRunPoints();
}


/**
 * Find the most recent head run submission timestamp.
 * 
 * @return {Date}  Headrun submission timestamp as Date object.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getLatestSubmissionTimestamp() {
  const sheet = ATTENDANCE_SHEET;
  const lastRow = sheet.getLastRow();
  const timestamp = sheet.getRange(lastRow, TIMESTAMP_COL).getValue();
  return new Date(timestamp);
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
  var sheet = LEDGER_SHEET;

  const TIMESTAMP_COL = 1;
  const HEAD_RUN_COL = 4;
  const ATTENDEES_COL = 5;
  const IS_ADDED_COL = 6;
  const NAMES_NOT_FOUND_COL = 7;

  const sheetImport = sheet.getSheetByName("Head Run Attendance");
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
  var sheet = LEDGER_SHEET;

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

  const pasteSheet = LEDGER_SHEET.getSheetByName(SHEET_NAME);
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
  const sheetPoints = LEDGER_SHEET.getSheetByName(SHEET_NAME);
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

