/*
Copyright 2024 Jikael Gagnon (for McGill Students Running Club)

Copyright 2024-25 Andrey Gonzalez (for McGill Students Running Club)

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

const LOG_TARGETS = {
  'id': LOG_INDEX.STRAVA_ACTIVITY_ID,      // (long)
  'name': LOG_INDEX.STRAVA_ACTIVITY_NAME,  // (string)
  'distance': LOG_INDEX.DISTANCE_STRAVA,   // meters (float)
  'moving_time': LOG_INDEX.MOVING_TIME,  // seconds (int)
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
 * @param {integer} [row = getValidLastRow_(LOG_SHEET)]  Target row.
 *                                                       Defaults to last valid row in `LOG_SHEET`.
 * 
 * @return {Object}  Strava activity.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 27, 2025
 * @update  May 28, 2025
 */
function findAndStoreStravaActivity(row = getValidLastRow_(LOG_SHEET)) {
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw Error("[PL] Please switch to the McRUN account before continuing");
  }

  // Check if Strava activity stored in sheet
  let activity = checkForExistingStrava_(row);
  if (activity) {
    Logger.log(`[PL] Strava activity found in log for row ${row}!`);
    return activity;
  }

  const timestamp = getSubmissionTimestamp_(row);
  const level = getRowLevel_(row);
  
  activity = popLevelRunFromExtras_(level);
  if (!activity || activity.length === 0) {
    // No activity stored, call Strava API instead
    // Get timestamp from row to use as filter
    const offset = 1000 * 60 * 60 * 2;    // 2 hours in seconds
    const limit = Math.floor((timestamp.getTime() + offset) / 1000);

    // Get all activities within timestamp range
    // For multiple activities, make educated guess and get by distance 
    const activities = getStravaStats_(timestamp, limit);
    activity = getActivityByLevel(level, activities);
  }

  if (!activity) {
    throw Error (`No Strava activity found for ${timestamp} (${level})`);
  }

  // Add mapUrl to activity if none found
  // Filename is timestamp and map store in Firebase storage
  if (!activity['mapUrl']) {
    activity = createAndAppendMap_(timestamp, activity);
  }

  // Set it to current row and return activity
  setStravaStats_(row, activity);
  return activity;
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
 * @date  Apr 1, 2025
 * @update  Apr 1, 2025
 */

function checkForExistingStrava_(row = getValidLastRow_(LOG_SHEET)) {
  const sheet = GET_LOG_SHEET_();
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


function getRowLevel_(row) {
  const sheet = GET_LOG_SHEET_();
  const eventTitle = sheet.getRange(row, LOG_INDEX.EVENT).getValue();

  const levelRegex = /\b(beginner|easy|intermediate|advanced)\b/i;
  const matches = eventTitle.match(levelRegex);
  return matches[1] || null;
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
 * @update  May 28, 2025
 */

function getStravaStats_(submissionTimestamp, toTimestamp) {
  // Get Unix Epoch value of timestamp to define search range
  const gracePeriod = 60 * 60 * 2   // In case the headrunner posted late
  const fromTimestamp = getUnixEpochTimestamp_(submissionTimestamp) - gracePeriod;

  // Get activity with time constraints
  const activities = getStravaActivity_(fromTimestamp, toTimestamp);
  if (!activities) {
    Logger.log(`[PL] No Strava activity has been found for the run that occured on ${submissionTimestamp}`);
  }

  return activities;
}


/**
 * Get Strava activity by level for multiple activities recorded at similar datetimes.
 * 
 * This helps sending the correct post-run email stats to attendee's level.
 * 
 * @param {string} level  Level of headrun (e.g. 'easy', 'intermediate').
 * @param {Object[]} activities  Array of Strava activities occurring at similar times.
 * @param {Object[]} levelHeadrunners  Array of headrunner objects for this level.
 * 
 * @return {Object|null}  Best-matching Strava activity, or null if none.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 27, 2025
 * @update  May 28, 2025
 */

function getActivityByLevel(level, activities, levelHeadrunners = []) {
  // Get level and if multiple activities, return activity by distance ascending.
  // E.g. activities = [{distance: 7km}, {distance: 3km}] -> Easy run = 3km, Intermediate = 7km
  // @see https://developers.strava.com/docs/reference/#api-models-DetailedActivity

  // I can also cross-reference with headrunner's Strava id and activities `athlete` property
  // Or with the activity's title `name` if level mentionned e.g. 'Easy Run!'
  if (!activities || activities.length === 0) return null;
  
  // // Convert levelHeadrunners to a set of athlete IDs for quick lookup
  // const headrunnerIds = new Set(levelHeadrunners.map(headrunner => headrunner.athleteId));
  
  // // Step 1: Try to match activities by athlete ID
  // let matchingActivities = activities.filter(act => headrunnerIds.has(act.athlete?.id));
  
  // if (matchingActivities.length === 0) {
  //   // Step 2: Fallback to all activities
  //   matchingActivities = [...activities];
  // }

  // Step 3: Sort activities by distance (in meters)
  const matchingActivities = activities;
  matchingActivities.sort((a, b) => a.distance - b.distance);

  // Find match according to distance, and store extras
  const match = getActivityByDistance(level, matchingActivities);
  const extraActivities = matchingActivities.filter(act => act !== match);
  extraActivities.length > 0 ? storeExtraActivities(extraActivities) : null;

  return match;

  /** Helper: assume that distance increases according to level */
  function getActivityByDistance(level, matchingActivities) {
    switch (level.toLowerCase()) {
      case 'beginner':
      case 'easy': return matchingActivities[0]; // Shortest distance
      case 'intermediate': return matchingActivities[Math.min(1, matchingActivities.length - 1)]; // Second shortest
      case 'advanced': return matchingActivities[matchingActivities.length - 1]; // Longest
      default: return matchingActivities[0] || null;
    }
  }

  // Save extra activities (excluding the selected one) in properties
  // Instead of calling Strava API multiple times
  function storeExtraActivities(extraActivities) {
    const scriptProps = PropertiesService.getScriptProperties();
    let currentExtras = scriptProps.getProperty(SCRIPT_PROPERTY_KEYS.extraStrava);

    // Combine current extras with input if applicable
    const toStore = currentExtras ? {...currentExtras, ...extraActivities} : extraActivities;
    scriptProps.setProperty(SCRIPT_PROPERTY_KEYS.extraStrava, JSON.stringify(toStore));
  }
}


/**
 * Remove a specific Strava activity from the extra activities stored for a level.
 * 
 * @param {string} level  The level of headrun (e.g., 'easy', 'intermediate').
 * @param {number} activityId  The Strava activity ID to remove.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date May 28, 2025
 */
function removeActivityFromExtra_(activityId) {
  if (!activityId) return;
  
  const scriptProps = PropertiesService.getScriptProperties();
  const propKey = SCRIPT_PROPERTY_KEYS.extraStrava;

  const jsonString = scriptProps.getProperty(propKey);
  if (!jsonString) return; // No stored activities
  
  try {
    const extraActivities = JSON.parse(jsonString);
    
    // Filter out the activity with the given ID
    const updatedActivities = extraActivities.filter(act => act.id !== activityId);
    
    // Update or remove the property if empty
    if (updatedActivities.length > 0) {
      scriptProps.setProperty(propKey, JSON.stringify(updatedActivities));
    } else {
      scriptProps.deleteProperty(propKey);
    }

  } catch (e) {
    Logger.log(`Error removing activity ID ${activityId}: ${e}`);

  }
}


/**
 * Retrieve extra Strava activities saved from previous API call.
 * 
 * @return {Object[]}  Extra Strava activities, or empty array if none found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date May 28, 2025
 */

function popLevelRunFromExtras_(level) {
  const extraActivities = getExtraActivities();
  const match = getActivityByLevel(level, extraActivities);

  // Clean up store and return match
  match ? removeActivityFromExtra_(match.id) : null;
  return match;

  function getExtraActivities() {
    const scriptProps = PropertiesService.getScriptProperties();
    const jsonString = scriptProps.getProperty(SCRIPT_PROPERTY_KEYS.extraStrava);
    return JSON.parse(jsonString) || [];
  }
}


function setStravaStats_(row, activity) {
  const sheet = GET_LOG_SHEET_();
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
  Logger.log(`[PL] Successfully imported Strava activity to row ${row} in Log Sheet!`);
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
