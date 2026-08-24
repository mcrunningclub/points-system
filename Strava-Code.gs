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

function runMe() {
  //const ROW = 72;
  //findAndStoreStravaActivity(ROW);
  //createMapForRow(ROW);

  Logger.log(`Now sending email to row ${ROW}...`);
  Utilities.sleep(4 * 1000)   // Grace period
  //sendStatsEmail(LOG_SHEET, ROW);
};

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
  const funcName = findAndStoreStravaActivity.name;
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw Error("[PL] Please switch to the McRUN account before continuing");
  }

  // Check if Strava activity stored in sheet
  let activity = checkExistingStravaActivity_(row);
  if (activity) {
    logAsPL_(`Strava activity found in log for row ${row}!`, funcName);
    return activity;
  }

  const timestamp = getTimestampInRow_(row);
  const level = getRowLevel_(row);
  
  activity = getMatchingActivityFromExtra_(level);
  if (!activity || Object.keys(activity).length === 0) {
    // No activity stored, call Strava API instead
    // Get timestamp from row to use as filter
    const offset = 1000 * 60 * 60 * 3;    // 3 hours in seconds
    const limit = Math.floor((timestamp.getTime() + offset) / 1000);

    // Get all activities within timestamp range
    // For multiple activities, make educated guess and get by distance 
    const activities = getStravaStats_(timestamp, limit);
    activity = getMatchingStravaActivity_(level, activities);

    if (!activity) {
      throw Error (`No Strava activity found for ${timestamp} (${level})`);
    }
  }

  // Successfully found Strava Activity
  logAsPL_(`Found a Strava activity to append to row #${row}!`, funcName);

  // Add mapUrl to activity if none found
  // Filename is timestamp and map store in Firebase storage
  if (!activity['mapUrl']) {
    activity = createAndAppendMap_(timestamp, activity);
    logAsPL_(`Created run map!`, funcName);
  }
  
  // Set it to current row and return activity
  setStravaStats_(row, activity);
  logAsPL_(`Set Strava stats for row #${row}!`, funcName);
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

function checkExistingStravaActivity_(row = getValidLastRow_(LOG_SHEET)) {
  const sheet = GET_LOG_SHEET();
  const startCol = LOG_COL.STRAVA_ACTIVITY_ID;
  const endCol = LOG_COL.MAP_URL;

  const stravaValues = sheet.getSheetValues(row, startCol, 1, endCol)[0];

  if (!stravaValues[0]) {
    return null;
  }

  const activityObj = {};
  const offset = LOG_COL.STRAVA_ACTIVITY_ID;

  for (const [id, index] of Object.entries(LOG_TARGETS)) {
    const relativeIndex = index - offset;
    activityObj[id] = stravaValues[relativeIndex];
  }

  return activityObj;
}


/**
 * Gets the headrun level for a given row in the log sheet.
 * 
 * @param {integer} row  Row of the activity.
 * 
 * @returns {string|null}  Level as string, e.g. "beginner". Null if not found.
 */
function getRowLevel_(row) {
  const sheet = GET_LOG_SHEET();
  const eventTitle = sheet.getRange(row, LOG_COL.EVENT).getValue();

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
  if (!activities || activities.length === 0) {
    logAsPL_(
      `No Strava activity has been found for the run that occured on ${submissionTimestamp}`,
      getStravaStats_.name
    );
  }

  return activities;
}


/**
 * Get Strava activity by level for multiple activities recorded at similar datetimes.
 * Try matching name of activity with level, or by distance otherwise.
 * 
 * This helps sending the correct post-run email stats to attendee's level.
 * 
 * @see https://developers.strava.com/docs/reference/#api-models-DetailedActivity
 * 
 * @param {string} level  Level of headrun (e.g. 'easy', 'intermediate').
 * @param {Object[]} activities  Array of Strava activities occurring at similar times.
 * 
 * @return {Object|null}  Best-matching Strava activity, or null if none.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 27, 2025
 * @update  Oct 5, 2025
 * 
 * ```js
 * const activities = [{name: 'Headrun Easy', distance: 7km}, {name: 'Morning run - Intermediate', distance: 3km}];
 * console.log(getActivityByLevel('Easy', activities))   // {name: 'Headrun Easy', distance: 7km}
 * ```
 */

function getMatchingStravaActivity_(level, activities) {
  if (!activities || activities.length === 0) return null;

  logAsPL_(`Now trying to match activity for level '${level}'...`, getMatchingStravaActivity_.name);
  console.log(activities);

  // First try to match by level in name, then try by distance
  const match = matchActivityByLevel(level, activities) ?? matchActivityByDistance(level, activities);

  // Store extra activities (if applicable)
  const extraActivities = activities.filter(act => act !== match);
  extraActivities.length > 0 ? storeExtraActivities(extraActivities) : console.log('No extras found!');

  return match;

  /** Helper 1: Strava activity contains level in name property */
  function matchActivityByLevel(level, matchingActivities){
    // If level contains 'easy' or 'beginner', search together because headrunners
    // sometimes label easy runs as beginner runs
    const regex = new RegExp(`${/beginner|easy/i.test(level) ?  "beginner|easy" : level}`, 'i');

    // Return activity if its name contains the target level, or undefined for no matches
    return matchingActivities.find(act => regex.test(act?.name));
  }


  /** Helper 2: assume that distance increases according to level */
  function matchActivityByDistance(level, matchingActivities) {
    // First sort by distance before matching
    matchingActivities.sort((a, b) => a.distance - b.distance);

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
    try {
      const toStore = currentExtras ? [...JSON.parse(currentExtras), ...extraActivities] : extraActivities;
      scriptProps.setProperty(SCRIPT_PROPERTY_KEYS.extraStrava, JSON.stringify(toStore));
    }
    catch (e) {
      console.error(`[PL#${storeExtraActivities.name}] Could not store extra activity. Please check manually`);
      console.error(`[PL#${storeExtraActivities.name}] Catch error: ${e.message}`);
    }
  }
}

/**
 * Remove a specific Strava activity from the extra activities stored for a level.
 * 
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
 * @param {string} level  Level of headrun (e.g. 'easy', 'intermediate').
 *
 * @return {Object[]}  Extra Strava activities, or empty array if none found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 28, 2025
 * @update  Sep 28, 2025
 */

function getMatchingActivityFromExtra_(level) {
  const extraActivities = getExtraActivities();
  logAsPL_(`Exited '${getExtraActivities.name}'`, getMatchingActivityFromExtra_.name);
  const match = getMatchingStravaActivity_(level, extraActivities);

  // Clean up store and return match
  match ? removeActivityFromExtra_(match.id) : console.log(`No match found :(`);
  return match;

  function getExtraActivities() {
    const scriptProps = PropertiesService.getScriptProperties();
    const jsonString = scriptProps.getProperty(SCRIPT_PROPERTY_KEYS.extraStrava);
    return JSON.parse(jsonString) || [];
  }
}


/**
 * Puts stats from specified Strava activity into the log sheet.
 * 
 * @param {number} row  Row to save the activity in.
 * @param {Object} activity  Strava activity to save stats from.
 */
function setStravaStats_(row, activity) {
  const sheet = GET_LOG_SHEET();
  const statsMap = Object.entries(LOG_TARGETS);

  // Get range from Strava Account to Map Polyline
  const startCol = LOG_COL.STRAVA_ACTIVITY_ID;
  const size = statsMap.length;
  const rangeToSet = sheet.getRange(row, startCol, 1, size);

  // Extract from activity and set in sheet
  const offset = size;
  const extracted = extractRunStats_(activity, statsMap, offset);
  rangeToSet.setValues([extracted]);

  // Log success mesage
  logAsPL_(`Successfully imported Strava activity to row ${row} in Log Sheet!`, setStravaStats_.name);
}


/**
 * Extract target run stats from Strava activity.
 * 
 * @see 'https://developers.strava.com/docs/reference/#api-models-SummaryActivity'
 * @see 'https://developers.strava.com/docs/reference/#api-models-ClubActivity'
 * 
 * @param {object} activity  A Strava object `SummaryActivity` or `ClubActivity`.
 * @param {{string: number}} statsMap  An object that maps stats to be extracted to the columns to save them to.
 * @param {number} offset  The amount to subtract from the column number. Default to 0.
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
