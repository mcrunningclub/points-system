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
    formattedAttendees = formattedAttendees.map(str => toTitleCase_(str));

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

function toTitleCase_(inputString) {
  return inputString.replace(/\w\S*/g, function (word) {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
  });
}


/**
 * Change the units in Strava activity to user-friendly values.
 * 
 * @param {Object} activity  Strava activity.
 * @return {Object}  Converted Strava activity in metric and US imperial values.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 30, 2025
 * @update  Apr 1, 2025
 */

function convertAndFormatStats_(activity) {
  const unitMap = getUnitsMap_();
  const formatMap = getNumberFormatMap_();

  // Duplicate properties of activity for both metric and imperial
  const converted = {metric : {...activity}, imperial : {...activity}};
  const systems = Object.keys(converted);

  Object.entries(activity).forEach(([key, value]) => {
    if (!(key in unitMap)) return;
    const units = convertAndFormat(key, value, unitMap, formatMap);
    systems.forEach(sys => converted[sys][key] = units[sys]);
  });

  return converted;

  /** Helper function to convert and format values according to their type mapping */
  function convertAndFormat(type, value, units, formats) {
    const factor = units[type];
    const format = formats[type];

    const operation = type === 'average_speed' ? (a, b) => a / b : (a, b) => a * b;

    return {
      metric: format(operation(factor.metric, value)),
      imperial: format(operation(factor.imperial, value)),
    };
  }
}


/** Returns the stat to unit conversion mapping in metric and imperial*/
function getUnitsMap_() {
  /** Metric Conversions  */
  const M_PER_SEC_TO_KM_PER_H = 3.6;
  const M_PER_SEC_TO_SEC_PER_KM = 1000;   // e.g. 1000 sec/km ÷ 4 m/s = 250 sec/km
  const M_TO_KM = 0.001;

  /** US Imperial Conversions  */
  const M_PER_SEC_TO_MI_PER_H = 2.237;
  const M_PER_SEC_TO_SEC_PER_MI = 1609;   // e.g. 1609 sec/mi ÷ 4 m/s = 402 sec/mi
  const M_TO_MI = 1 / 1609;
  const M_TO_FEET = 3.2808;

  return {
    'distance': pack(M_TO_KM, M_TO_MI),
    'elapsed_time': pack(1, 1),   // Leave as seconds to format as 'mm:ss' later
    'average_speed': pack(M_PER_SEC_TO_SEC_PER_KM, M_PER_SEC_TO_SEC_PER_MI),  // Likewise
    'max_speed': pack(M_PER_SEC_TO_KM_PER_H, M_PER_SEC_TO_MI_PER_H),
    'total_elevation_gain' : pack(1, M_TO_FEET),
  }

  function pack(aMetric, aImperial) {
    return {metric : aMetric, imperial : aImperial};
  }
}


function getNumberFormatMap_() {
  return {
    'distance': x => toFixedTruncate(x, 2),
    'elapsed_time': x => toMinuteSeconds(x),
    'average_speed': x => toMinuteSeconds(x),
    'max_speed': x => x.toFixed(1),
    'total_elevation_gain' : x => {
      const sign = (x > 0) ? '+' : '';
      return `${sign}${x.toFixed(0)}`;
    },
  }

  /** Replaced .toFixed() to improve accuracy, e.g. 5.9989 -> 5.99 instead of 6.00 */
  function toFixedTruncate(num, digits) {
    const factor = Math.pow(10, digits);
    const truncated = Math.floor(num * factor) / factor;

    return truncated.toFixed(digits);  // Convert to string and pad with zeros
  }

  /** Format duration as 'mm:ss' */
  function toMinuteSeconds(t) {
    const totalMin = Math.floor(t / 60);
    const totalSec = `${Math.round(t % 60)}`;

    return totalMin + ':' + totalSec.padStart(2, "0");
  }
}


function testActivityFormatting() {
  const activity = {
    'distance': 5998.9,
    'elapsed_time': 2766,
    'average_speed': 3.616,
    'max_speed': 4.84,
    'total_elevation_gain' : 24.7,
    'points' : 100,
  }
  
  const ret = convertAndFormatStats_(activity);
  console.log(activity);
  console.log(ret);
}