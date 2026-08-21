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
 * @update  May 19, 2025
 */

function sortLogsByTimestamp() {
  const sheet = LOG_SHEET;

  // Sort timestamps in ascending order, without the header row
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort(3);
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
 * Change the units in Strava activity to user-friendly values and format them.
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
  // Duplicate properties of activity for both metric and imperial
  const metricAcivity = { ...activity };
  const imperialActivity = { ...activity };

  for (const [key, value] of activity) {
    if (key in UNIT_MAP) {
      const convertFactor = UNIT_MAP[key];
      const formatFunc = NUMBER_FORMAT_MAP[key];
      metricAcivity[key] = formatFunc(convertFactor.metric * value);
      imperialActivity[key] = formatFunc(convertFactor.imperial * value);
    }
  }

  return {
    metric: metricActivity, 
    imperial: imperialActivity
  };
}

/**
 * Truncate decimal number to given number of digits.
 * 
 * Replaced .toFixed() to improve accuracy, e.g. 5.9989 -> 5.99 instead of 6.00
 * 
 * @param {float} num  The number to truncate
 * @param {integer} digits  Number of decimal places to keep
 * @return {float}  Truncated number
 */
function toFixedTruncate(num, digits) {
  const factor = Math.pow(10, digits);
  const truncated = Math.floor(num * factor) / factor;

  return truncated.toFixed(digits);  // Convert to string and pad with zeros
}

/** 
 * Format duration as 'mm:ss'.
 * 
 * @param {number} t  Duration in seconds
 * @return {string}  Duration in format mm:ss
 */
function toMinuteSeconds(t) {
  const totalMin = Math.floor(t / 60);
  const totalSec = `${Math.round(t % 60)}`;

  return totalMin + ':' + totalSec.padStart(2, "0");
}