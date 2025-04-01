/*
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
 * Sorts log sheet by event timestamp ascending.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 28, 2023
 * @update  Mar 31, 2025
 */

function sortTimestampByAscending() {
  const sheet = LOG_SHEET;

  // Sort timestamps in ascending order, without the header row
  const range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.sort(3);
}


/**
 * Format specific columns in `Head Run Attendance`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2023
 * @update  Mar 31, 2025
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
  let attendeesCell = sheet.getRange(getValidLastRow(sheet), 5);

  if (attendeesCell.toString().length > 1) {
    let splitArray = attendeesCell.getValue().split('\n');  // split the string into an array;
    if (splitArray.length < 2) splitArray = attendeesCell.getValue().split(',');  // split the string into an array;

    // Trim whitespace from strings and set to Title Case
    let formattedAttendees = splitArray.map(str => str.trim());
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
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 30, 2023
 * @update  Oct 30, 2023
 */

function toTitleCase(inputString) {
  return inputString.replace(/\w\S*/g, function (word) {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
  });
}


/**
 * Change the units from Strava activity to user-friendly units.
 * 
 * @param {Object} activity  Strava activity.
 * @param {Boolean} isMetric  True if metric system is used, else imperial system.
 * @return {Object}  Converted Strava activity.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 30, 2025
 * @update  Mar 31, 2025
 */

function convertAndFormatStat(activity, isMetric) {
  const units = getUnitsMap_(isMetric);
  const formats = getNumberFormatMap(); 

  Object.entries(activity).forEach(([key, value]) => {
    if (!(key in units)) return;
    const convertedVal = convertUnit_(key, value);
    activity[key] = formatNumber(key, convertedVal);
  });

  function convertUnit_(type, value) {
    if (type === 'average_speed') {
      return units[type] / value;
    }
    return value * units[type];
  };

  function formatNumber(key, value) {
    return formats[key](value);
  }
}

/** Returns the stat to unit conversion mapping in metric or imperial*/
function getUnitsMap_(isMetric) {
  const SEC_TO_MIN = 1 / 60;

  /** Metric Conversions  */
  const M_PER_SEC_TO_KM_TO_H = 3.6;
  const M_PER_SEC_TO_KM_PER_MIN = 100 / 6;
  const M_TO_KM = 0.001;

  /** US Imperial Conversions  */
  const M_PER_SEC_TO_MILES_TO_H = 2.237;
  const M_PER_SEC_TO_MILES_PER_MIN = 26.822;
  const M_TO_MILES = 1 / 1609;
  const M_TO_FEET = 3.2808;

  return {
    'distance': isMetric ? M_TO_KM : M_TO_MILES,
    'elapsed_time': SEC_TO_MIN,
    'average_speed': isMetric ? M_PER_SEC_TO_KM_PER_MIN : M_PER_SEC_TO_MILES_PER_MIN,
    'max_speed': isMetric ? M_PER_SEC_TO_KM_TO_H : M_PER_SEC_TO_MILES_TO_H,
    'total_elevation_gain' : isMetric ? 1 : M_TO_FEET,
  }
}


function getNumberFormatMap() {
  return {
    'distance': x => x.toFixed(2),
    'elapsed_time': x => x.toFixed(0),
    'average_speed': x => x.toFixed(2),
    'max_speed': x => x.toFixed(1),
    'total_elevation_gain' : x => {
      const sign = (x > 0) ? '+' : '';
      return `${sign}${x.toFixed(0)}`;
    },
  }
}

function test() {
  const activity = {
    'distance': 5830.3,
    'elapsed_time': 2638,
    'average_speed': 3.013,
    'max_speed': 4.54,
    'total_elevation_gain' : -21.4,
  }
  
  convertAndFormatStat(activity, true)
  console.log(activity);
}