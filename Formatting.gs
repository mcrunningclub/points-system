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


/** 
 * Sorts sheet by first name ascending.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 28, 2023
 * @update  Nov 28, 2023
 */

function sortNameByAscending() {
  var sheet = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);
  
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
  
  // Sorts values by the `First Name` column in ascending order
  range.sort(3);
  return;
}


/**
 * Format specific columns in `Head Run Attendance`.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2023
 * @update  Oct 30, 2023
 */

function formatSpecificColumns() {
  const sheet = LEDGER_SS.getSheetByName("Head Run Attendance");
  
  const rangeListToBold = sheet.getRangeList(['A2:A', 'D2:D']);
  rangeListToBold.setFontWeight('bold');  // Set ranges to bold

  const rangeListToWrap = sheet.getRangeList(['B2:E', 'G2:H']);
  rangeListToWrap.setWrap(true);  // Turn on wrap

  const rangeAttendees = sheet.getRange('E2:E');
  rangeAttendees.setFontSize(9);  // Reduce font size for `Attendees` column

  const rangeIsAdded = sheet.getRange(2, 6, 21);
  rangeIsAdded.insertCheckboxes();  // Add checkbox

  // Formats only last row of attendees
  var attendeesCell = sheet.getRange(getValidLastRow(sheet), 5);
  
  if(attendeesCell.toString().length > 1) {
    var splitArray = attendeesCell.getValue().split('\n');  // split the string into an array;
    if (splitArray.length < 2) splitArray = attendeesCell.getValue().split(',');  // split the string into an array;

    // Trim whitespace from strings and set to Title Case
    var formattedAttendees = splitArray.map(str => str.trim());
    formattedAttendees = formattedAttendees.map(str => toTitleCase(str));
    
    var newValue = formattedAttendees.join('\n');       // combine all array elements into single string
    attendeesCell.setValue(newValue);
  }
}


/**
 * Formats string to Title Case.
 * 
 * @param {string} inputString  String to format.
 * @return {string}  String in title case.
 * 
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2023
 * @update  Oct 30, 2023
 */

function toTitleCase(inputString) {
  return inputString.replace(/\w\S*/g, function(word) {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
  });
}
