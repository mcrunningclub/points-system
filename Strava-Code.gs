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

const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/'
const ACTIVITIES_ENDPOINT = 'athlete/activities'
const MAPS_FOLDER = 'run_maps'

// SCRIPT PROPERTIES (MAKE SURE THAT NAMES MATCHES ACTUAL STORE)
const SCRIPT_PROPERTY_KEYS = {
  clientID: 'CLIENT_ID',
  clientSecret: 'CLIENT_SECRET',
  googleMapAPI: 'GOOGLE_MAPS_API_KEY',
};

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
 * Strava API playground for developer.
 */

function stravaPlayground() {
  /** Club member example */
  function runExample(label) {
    switch (label) {
      case 'A': {
        /** Club activites example */
        var endpoint = 'clubs/693906/members'
        var query_object = { "include_all_efforts": true };
        return callStravaAPI_(endpoint, query_object);

      }

      case 'B': {
        /** Club activites example */
        var endpoint = 'activities/13889807290';
        var response = callStravaAPI_(endpoint, { "include_all_efforts": true });
        return response["map"]["polyline"];
      }

      case 'C': {
        /** Individual athlete example */
        var endpoint = 'athletes/13968699602/stats' // 'athlete/activities';
        var response = callStravaAPI_(endpoint, {})[0];
        getMapBlob_(response['map']['summary_polyline'], 'example.png');
      }
    }
  }

  // Choose which function to run and log response
  const response = runExample('E');
  console.log('Result: ' + response);
}


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
 * @update  Apr 1, 2025
 */

function findAndStoreStravaActivity(row = getValidLastRow(LOG_SHEET)) {
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw Error ("Please switch to the McRUN account before continuing");
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
    const filename = Utilities.formatDate(timestamp, TIMEZONE, 'EEE-d-MMM-yyyy-hh:mm');
    activity['mapUrl'] = createStravaMap_(activity, filename);
    setStravaStats_(row, activity);
  }

  return activity;
}


function getAllActivities() {
  const startRow = 88;
  const endRow = 88; //getValidLastRow(LOG_SHEET);
  for(let row = startRow; row <= endRow; row++) {
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

  if(!stravaValues[0]) {
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


function createStravaMap_(activity, name) {
  // Extract polyline and save headrun route as map
  const polyline = activity['map']['polyline'] ?? activity['map']['summary_polyline'];

  if (polyline) {
    const response = saveMapForRun_(polyline, name).getHeaders();
    
    // Get file by id or name, then set permission to allow downloading
    const file = response['file_id'] ? getFileById_(response['file_id']) : getFileByName_(name);
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    // Return download url from file
    return file.getUrl();
  }
  
  return '';
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


/**
 * Save polyline as image using Google Map API and Make.com automation.
 * 
 * @param {string} polyline  Encoded Google Map polyline string.
 * @param {string} name  Name for map.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 27, 2025
 * @update  Apr 2, 2025
 */

function saveMapForRun_(polyline, name) {
  const postUrl = buildPostUrl_(polyline, "580x420");
  return postToMakeWebhook_(postUrl, name);
}

/** Helper 1: Construct postUrl for Make webhook */
function buildPostUrl_(polyline, imgSize = "580x420") {
  const propertyStore = PropertiesService.getScriptProperties();
  const apiKey = propertyStore.getProperty(SCRIPT_PROPERTY_KEYS.googleMapAPI); // Replace with your API Key

  const googleCloudMapId = 'bfeadd271a2b0a58';  //'2ff6c54f4dd84b16';
  const pathColor = '0xEB4E3D';

  const queryObj = {
    size: imgSize,
    map_id: googleCloudMapId,
    key: apiKey,
    //path: `color:${pathColor}` + '|' + `enc:${polyline}`,
    path: `enc:${polyline}`,
  }

  return MAPS_BASE_URL + queryObjToString_(queryObj);
}

/** Helper 2: Call Make webhook */
function postToMakeWebhook_(postUrl, mapName) {
  const webhookUrl = "https://hook.us1.make.com/8obb3hb6bzwgi7s4nyi8yfghb3kxsksc";
  const payload = JSON.stringify({ url: postUrl, name: mapName });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  Logger.log("Make Webhook Response: " + response.getContentText());
  return response;
}


/**
 * Save the map of a Strava activity as `MAPS_FOLDER/<Unix Epoch timestamp>.png`
 * 
 * @deprecated  Use Make Webhook instead.
 * 
 * @param {string} polyline  A encoded polyline representing a path.
 * @param {integer} timestamp  Name to save file as.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Dec 1, 2024
 * @update  Mar 27, 2025
 */

function saveMapAsBlob_(polyline, timestamp) {
  if (!polyline) {
    return Logger.log('Map cannot be created: no polyline found for this activity');
  }

  const runMap = Maps.newStaticMap()
    .setSize(800, 600) // Adjust size as needed
    .setFormat(Maps.StaticMap.Format.PNG)
    .setMapType(Maps.StaticMap.Type.ROADMAP) // Use ROADMAP for a minimalist style
    .setPathStyle(3, "0x000000FF", "0x00000000") // Thin black route line, transparent fill
    .addPath(polyline)
  ;

  // Get save locati on using timestamp
  const saveLocation = getSaveLocation(timestamp);

  // Save runMap as as image to specified location
  const mapBlob = Utilities.newBlob(runMap.getMapImage(), 'image/png', saveLocation);
  const file = DriveApp.createFile(mapBlob);

  // Display success message
  Logger.log(`Successfully saved map as ${saveLocation}`);

  // Set permission to allow downloading
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  return file.getDownloadUrl();

  /** Helper function to get location of file to save */
  function getSaveLocation(submissionTime) {
    return MAPS_FOLDER + '/' + submissionTime.toString() + '.png'
  }
}
