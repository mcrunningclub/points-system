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

const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/';
const ACTIVITIES_ENDPOINT = 'athlete/activities';

const LOG_TARGETS = {
  'id': LOG_INDEX.STRAVA_ACTIVITY_ID,      // (long)
  'name': LOG_INDEX.STRAVA_ACTIVITY_NAME,  // (string)
  'distance': LOG_INDEX.DISTANCE_STRAVA,   // meters (float)
  'elapsed_time': LOG_INDEX.ELAPSED_TIME,  // seconds (int)
  'average_speed': LOG_INDEX.PACE,         // m per sec (float)
  'max_speed': LOG_INDEX.MAX_SPEED,        // m per sec (float)
  'total_elevation_gain': LOG_INDEX.ELEVATION,   // meters (float)
  'map': LOG_INDEX.MAP_POLYLINE,
  'mapUrl': LOG_INDEX.MAP_URL,
};

// Simple logging of multi-line message. Improves readability in code.
const prettyLog_ = (...msg) => console.log(msg.join('\n'));


/**
 * Return Strava activity in `row`. If Strava activity not found in `LOG_SHEET`,
 * call Strava API using `timestamp` as searching target.
 * 
 * 
 * @param {integer} [row = getValidLastRow(LOG_SHEET)]  Target row.
 *                                                      Defaults to last valid row in `LOG_SHEET`.
 * 
 * @return {Object}  Strava activity.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 27, 2025
 * @update  Apr 4, 2025
 */

function findAndStoreStravaActivity(row = getValidLastRow(LOG_SHEET)) {
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw Error("Please switch to the McRUN account before continuing");
  }

  // Check if Strava activity stored
  let activity = checkForExistingStrava_(row);
  if (activity) {
    Logger.log(`Strava activity found in log for row ${row}!`);
    return activity;
  }

  // No activity stored, call Strava API instead
  // Get timestamp from row
  const timestamp = getSubmissionTimestamp(row);
  const offset = 1000 * 60 * 60 * 2;    // 2 hours in seconds
  const limit = Math.floor((timestamp.getTime() + offset) / 1000);

  // Save stats to log sheet and store map to Drive.
  // Filename is timestamp. Download url added to `activity` obj
  activity = getStravaStats_(timestamp, limit);

  // If activity available, add mapUrl and save in log sheet
  if (activity) {
    const formattedTS = Utilities.formatDate(timestamp, TIMEZONE, "EEE-d-MMM-yyyy-k\'h\'mm");
    const filename = `headrun-map-${formattedTS}.png`
    const mapBlob = createStravaMap_(activity, filename);

    // Upload image to Google Cloud Storage and get sharing link
    activity['mapUrl'] = uploadImageToBucket(mapBlob, filename);
    setStravaStats_(row, activity);
  }

  return activity;
}


function getAllActivities() {
  const startRow = 88;
  const endRow = 88; //getValidLastRow(LOG_SHEET);
  for (let row = startRow; row <= endRow; row++) {
    findAndStoreStravaActivity(row);
  }
}


/**
 * Verify if Strava activity already stored in log.
 * 
 * Prevents redundant Strava API call.
 * 
 * @param {integer} [row = getValidLastRow(LOG_SHEET)]  Target row.
 *                                                      Defaults to last valid row in `LOG_SHEET`.
 * 
 * @return {Object}  Previously stored Strava activity.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Apr 1, 2025
 * @update  Apr 1, 2025
 */

function checkForExistingStrava_(row = getValidLastRow(LOG_SHEET)) {
  const sheet = LOG_SHEET;
  const startCol = LOG_INDEX.STRAVA_ACTIVITY_ID;
  const endCol = LOG_INDEX.MAP_URL;

  const stravaValues = sheet.getSheetValues(row, startCol, 1, endCol)[0];

  if (!stravaValues[0]) {
    return null;
  }

  const activityObj = {};
  const offset = LOG_INDEX.STRAVA_ACTIVITY_ID;

  for (const [id, index] of Object.entries(LOG_TARGETS)) {
    const relativeIndex = index - offset;
    activityObj[id] = stravaValues[relativeIndex];
  }

  return activityObj;
}


/**
 * Get Strava activity of most recent head run submission.
 * 
 * @param {Date} submissionTimestamp  Date representation of headrun timestamp.
 * @param {integer} toTimestamp  Max timestamp for map search in seconds.
 * 
 * @return {Object}  Strava activity with appended mapUrl
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 27, 2025
 * @update  April 1, 2025
 */

function getStravaStats_(submissionTimestamp, toTimestamp) {
  // Get Unix Epoch value of timestamp to define search range
  const gracePeriod = 60 * 60 * 2   // In case the headrunner posted late
  const fromTimestamp = getUnixEpochTimestamp_(submissionTimestamp) - gracePeriod;

  // Get activity with time constraints
  const activity = getStravaActivity_(fromTimestamp, toTimestamp);

  if (!activity) {
    Logger.log(`No Strava activity has been found for the run that occured on ${submissionTimestamp}`);
  }

  return activity;
}


function setStravaStats_(row, activity) {
  const sheet = LOG_SHEET;
  const statsMap = Object.entries(LOG_TARGETS);

  // Get range from Strava Account to Map Polyline
  const startCol = LOG_INDEX.STRAVA_ACTIVITY_ID;
  const size = statsMap.length;
  const rangeToSet = sheet.getRange(row, startCol, 1, size);

  // Extract from activity and set in sheet
  const offset = size - 1;
  const extracted = extractRunStats_(activity, statsMap, offset);
  rangeToSet.setValues([extracted]);

  // Log success mesage
  Logger.log(`Successfully imported Strava activity to row ${row} in Log Sheet!`);
}


/**
 * Convert a Date timestamp to a Unix Epoch timestamp.
 * 
 * @param {Date} timestamp  Timestamp to convert.
 * @return {integer}  Number of seconds elapsed since January 1, 1970.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getUnixEpochTimestamp_(timestamp) {
  return Math.floor(timestamp.getTime() / 1000);
}


/**
 * Extract target run stats from Strava activity.
 * 
 * @see 'https://developers.strava.com/docs/reference/#api-models-SummaryActivity'
 * @see 'https://developers.strava.com/docs/reference/#api-models-ClubActivity'
 * 
 * @param {object} activity  A Strava object `SummaryActivity` or `ClubActivity`.
 * @return {object}  Extracted stats from `activity`.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 22, 2025
 * @update  Mar 27, 2025
 */

function extractRunStats_(activity, statsMap, offset = 0) {
  const valArr = [];
  for (const [stat, index] of statsMap) {
    const relativeIndex = index - offset;
    valArr[relativeIndex] = activity[stat] || "";
  }

  return valArr;
}
