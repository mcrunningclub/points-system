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
  'athlete' : LOG_INDEX.STRAVA_ACCOUNT,
  'name' : LOG_INDEX.STRAVA_ACTIVITY_NAME,
  'distance' : LOG_INDEX.DISTANCE_STRAVA,
  'elapsed_time' : LOG_INDEX.ELAPSED_TIME,
  'average_speed' : LOG_INDEX.PACE,
  'max_speed' : LOG_INDEX.MAX_SPEED,
  'total_elevation_gain' : LOG_INDEX.ELEVATION,
  'map' : LOG_INDEX.MAP_POLYLINE,
  'mapUrl' : LOG_INDEX.MAP_URL,
};

// Simple logging of multi-line message. Improves readability in code.
const prettyLog_ = (...msg) => console.log(msg.join('\n'));


/**
 * Strava API playground for user.
 */

function stravaPlayground() {
  /** Club member example */
  function runExample(label) {
    switch(label) {
      case 'A' : {
        /** Club activites example */
        var endpoint = 'clubs/693906/members'
        var query_object = {"include_all_efforts" : true};
        return callStravaAPI_(endpoint, query_object);
        
      }

      case 'B' : {
        /** Club activites example */
        var endpoint = 'activities/13889807290';
        var response = callStravaAPI_(endpoint, {"include_all_efforts" : true});

        return response["map"]["polyline"];
        //return extractRunStats_(response[0]);
      }

      case 'C' : {
        /** Individual athlete example */
        var endpoint = 'athletes/13968699602/stats' // 'athlete/activities';
        var response = callStravaAPI_(endpoint, {})[0];
        getMapBlob(response['map']['summary_polyline'], 'example.png');
        return extractRunStats_(response);
      }

      case 'D' : {
        /** Activity tagged by headrunner */
        const timestamp = new Date('2025-02-22 10:00:00');
        const upperLimit = 3600 * 1000;   // 1 hour in milliseconds
        const upperTimestamp = new Date(timestamp.getTime() + upperLimit);
        return saveMapForRun_(timestamp, upperTimestamp);
      }
    }
  }
  
  // Choose which function to run and log response
  const response = runExample('D');
  console.log('Result: ' + response);
}


function saveStravaStats_(row, submissionTimestamp, maxDate = new Date()) {
  // Get Unix Epoch value of timestamps to define search range
  const toTimestamp = getUnixEpochTimestamp_(maxDate);
  const fromTimestamp = getUnixEpochTimestamp_(submissionTimestamp);

  const activity = getStravaActivity(fromTimestamp, toTimestamp);

  // Extract polyline and save path as map
  //const polyline = activity['map']['summary_polyline'];
  //const mapDownloadUrl = saveMapAsFile_(polyline, fromTimestamp);
  const mapDownloadUrl = 'https://drive.google.com/uc?id=1GET38bgwPRuWXWI7GLgjtHztghE2K7mT&export=download';

  // Add map url to activity
  activity['mapUrl'] = mapDownloadUrl;  // Verify property name in LOG_TARGETS

  // Save activity to row in sheet
  setStravaStats_(row, activity);
  Logger.log('Successfully imported Strava activity to Log Sheet!');
}


function setStravaStats_(row, activity) {
  const sheet = LOG_SHEET;
  const statsMap = LOG_TARGETS;
  const statsArr = Object.entries(statsMap);

  // Get range from Strava Account to Map Polyline
  const startCol = LOG_INDEX.STRAVA_ACCOUNT;
  const size = statsArr.length;
  const rangeToSet = sheet.getRange(row, startCol, 1, size);

  // Extract from activity and set in sheet
  const offset = size - 1;
  const extracted = extractRunStats_(activity, statsArr, offset);
  rangeToSet.setValues([extracted]);
}


function getStravaActivity(fromTimestamp, toTimestamp) {
  // Package query for Strava API
  const queryObj = { 'after': fromTimestamp, 'before': toTimestamp };

  const endpoint = ACTIVITIES_ENDPOINT;
  const response = callStravaAPI_(endpoint, queryObj);

  if (response.length === 0) {
    // Create an instance of ExecutionError with a custom message
    const errorMessage = `No Strava activity has been found for the run that occured on ${new Date(fromTimestamp)}`;
    throw new Error(errorMessage);
  }

  const activity = response[0];  // Assume first activity is the target
  return activity;
}


/**
 * Makes an API request to the given endpoint with the given query.
 * 
 * Inspired by original function `run` in `apps-script-oauth2/samples/Strava.gs`
 * 
 * @param {string} endpoint  Strava API endpoint.
 * 
 * @param {object} [query_object = {}]  Param-value pair.
 *                                      Defaults to empty object.
 * 
 * @return {string}  Response of API call.
 * 
 * ### Sample script
 * ```javascript
 * const endpoint = 'clubs/693906/activities';
 * const queryObj = {"param1": val1, "param2": val2};
 * const response = callStravaAPI(endpoint, queryObj);
 * ```
 * 
 * ### Info
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 7, 2024
 * @update  Mar 23, 2025
 */

function callStravaAPI_(endpoint, query_object = {}) {
  // Set up the service
  const service = getStravaService_();

  // Verify is access authorized already
  if (!service.hasAccess()) {
    // Display authorization url and exit
    return prettyLog_(
      'App has no access yet.',
      'Open the following URL to gain authorization from Strava and re-run the script.',
      service.getAuthorizationUrl()
    ); 
  }

  // Authorization completed.
  Logger.log('App has access.');
  
  // Get API endpoint
  endpoint = STRAVA_BASE_URL + endpoint;
  const query_string = queryObjToString_(query_object);

  const headers = {
    Authorization: 'Bearer ' + service.getAccessToken()
  };

  const options = {
    headers: headers,
    method: 'GET',
    muteHttpExceptions: false,
  };

  // Return Strava API response
  const urlString = endpoint + query_string;
  return JSON.parse(UrlFetchApp.fetch(urlString, options));
}


/**
 * Maps an Object containing param-value pairs to a query string.
 *  
 * @param {object} [query_object = {}]  Param-value pair.
 *                                      Defaults to empty object.
 * 
 * @return {string}  String value of query object.
 * 
 * ### Sample script
 * ```javascript
 * const queryObj = {"param1": val1, "param2": val2};
 * const ret = queryObjToString(queryObj);
 * Logger.log(ret)  // "?param1=val1&param2=val2"
 * ```
 * 
 * ### Info
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 7, 2024
 * @update  Mar 22, 2025
 */

function queryObjToString_(query_object = {}) {
  // Check if object empty
  if (query_object.length === 0) {
    return '';
  }
  const param_value_list = Object.entries(query_object);
  const param_strings = param_value_list.map(([param, value]) => `${param}=${value}`);
  const query_string = param_strings.join('&');
  return '?' + query_string;
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
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
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


function saveMapForLatestRun() {
  const submissionTimestamp = getLatestSubmissionTimestamp();
  saveMapForRun_(submissionTimestamp);
}


/**
 * Takes a Strava API response for a given activity and saves an
 * image of the map to the desired location.
 * 
 * @param {string} polyline  A encoded polyline representing a path.
 * @param {string} filename  Name to save file as.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Dec 1, 2024
 * @update  Mar 25, 2025
 */

function getMapBlob(polyline, filename) {
  if (!polyline) {
    return Logger.log('Map cannot be created: no polyline found for this activity');
  }
  
  const runMap = Maps.newStaticMap();
  runMap.addPath(polyline);

  // Save runMap as as image to specified location
  const mapBlob = Utilities.newBlob(runMap.getMapImage(), 'image/png', filename);
  
  // Display success message
  Logger.log(`Successfully created map blob as ${filename}.png`);
  return mapBlob;
}


/**
 * Takes a Strava API response for a given activity and saves an
 * image of the map to the desired location.
 * 
 * @deprecated  Use Make Webhook instead.
 * 
 * @param {string} polyline  A encoded polyline representing a path.
 * @param {integer} timestamp  Name to save file as.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Dec 1, 2024
 * @update  Mar 27, 2025
 */

function saveMapAsFile_(polyline, timestamp) {
  if (!polyline) {
    return Logger.log('Map cannot be created: no polyline found for this activity');
  }

  const runMap = Maps.setAuthentication().newStaticMap()
   .setMapType(Maps.StaticMap.Type.ROADMAP)
   .setPathStyle(6, Maps.StaticMap.Color.RED, "0x00000000")
   .addPath(polyline)
  ;

  // Get save locati on using timestamp
  const saveLocation = getSaveLocation(timestamp);

  // Save runMap as as image to specified location
  const mapBlob = Utilities.newBlob(runMap.getMapImage(), 'image/png', saveLocation)
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


/**
 * Get Strava activity of most recent head run submission.
 * 
 * Save the map as `MAPS_FOLDER/<Unix Epoch timestamp of submisstion>.png`
 * 
 * @param {Date} submissionTimestamp  Date representation of headrun timestamp.
 * @param {Date} maxDate  Max timestamp for map search.
 *
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Dec 1, 2024
 * @update  Mar 27, 2025
 */

function saveMapForRun_(submissionTimestamp, maxDate = new Date()) {  
  // Get string representation of timestamp
  const submissionTimestampStr = submissionTimestamp.toString();

  // Get Unix Epoch value of timestamps to define search range
  const toEpochTime = getUnixEpochTimestamp_(maxDate);
  const fromEpochTime = getUnixEpochTimestamp_(submissionTimestamp);

  const queryObj = { 'after': fromEpochTime, 'before': toEpochTime };

  const endpoint = ACTIVITIES_ENDPOINT;
  const response = callStravaAPI_(endpoint, queryObj);

  if (response.length === 0) {
    // Create an instance of ExecutionError with a custom message
    const errorMessage = `No Strava activity has been found for the run that occured on ${submissionTimestampStr}`;
    throw new Error(errorMessage);
  }

  const activity = response[0];  // Assume first activity is the target

  // Extract polyline and save path as map
  const polyline = activity['map']['summary_polyline'];
  postToMakeWebhook_(polyline);

  return Logger.log(polyline);

  saveThisPolyline(polyline);
}


/**
 * Example scripts to save routes using Google Maps API
 */

function createMinimalistRouteMap_() {
  const propertyStore = PropertiesService.getScriptProperties();
  const apiKey = propertyStore.getProperty(myScriptKeys.googleMapAPI); // Replace with your API Key
  
  const encodedPolyline = "irwtGh~a`MG?Y]KSW[QMMOUAQI{ALQDK?OGWc@e@[OSUOMQYQu@u@GC[_@OIs@i@QKEKOIQ_@GCSUE?]WOUKEGMc@[m@s@ECG?UOGKSSc@[GKOIYYIEGKQEcA{@QCOQKIE@MKOUa@WW]QI[WaAcAEAm@e@MGw@w@E?QOYYE?IMOGAGQEWSY_@[QsAkAWKKQSKCGUS_@m@Y?SSGBENMVY~@I@ECMSUOIIQKGMq@m@SWKAUOUUGKKEY[WMCGUSIMI?_@MGMUOCEWOCIe@k@OYQMCG]YOIGMOIISo@]CGm@k@e@a@SK[]OKEK?GR[@ITs@DCj@aBBCFWZu@ZgADCDWN]Jq@CEOCWa@KACCOOOKMSQI{@y@MKYKa@G_@KUAUEy@SQCUK]CIG}@Io@Uw@MSKaBYIEc@GeBc@I?a@MWA_Aa@I?WGE?QC[IW@M`@E`@OTUp@J_@LSL_@Uv@cAhCOl@_A~BCD_@lAQ\Oj@S`@CJILAHSh@CPWh@W|@k@hACJIHYhAGHKZKLKZBFPHLHb@d@HJ@LNFTTRLH?HNNHd@f@NH\`@FBNNB@VZNHr@p@v@x@t@n@Td@TTRP\PJNd@`@JPp@n@ZTFNv@p@LR\VFJh@d@FH`@^^d@Nf@F@FCFKN@LId@_BNe@R_@Pq@Ti@X_ATc@H]P_@`A}C\q@F]l@yARu@NWBWj@sALe@LULe@DANVNBLFNLPVf@RrL~KFJp@f@^`@RNJLJD^^FEFY^}@JQFCRPV\XTLPb@RRX\Tt@p@HBj@f@LRB@\XD?PRRJj@f@ZRNPTLNPRL@Dj@f@f@VPRVLf@h@ZT\^NHTTTFb@j@F?FDLRPBHNd@Vn@f@XNBLBDXL@Df@`@Z\h@b@DJb@RD?RZLNt@^^b@VV^h@v@r@ZZp@j@j@j@NNFRJLVEh@BX?t@Sp@`@VLPR^^";

  // Create a new static map
  const map = Maps.newStaticMap()
    .setSize(800, 600) // Adjust size as needed
    .setFormat(Maps.StaticMap.Format.PNG)
    .setMapType(Maps.StaticMap.Type.ROADMAP) // Use ROADMAP for a minimalist style
    .setPathStyle(3, "0x000000FF", "0x00000000") // Thin black route line, transparent fill
    .addPath(encodedPolyline);

  // Get the URL of the generated static map
  const mapUrl = `${map.getMapUrl()}&key=${apiKey}`;
  Logger.log("Map URL: " + mapUrl);

  return mapUrl;
}


function saveMapToDrive() {
  var mapUrl = createMinimalistRouteMap_();
  var response = UrlFetchApp.fetch(mapUrl);
  var blob = response.getBlob().setName("route-map2.png");

  var file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const imageUrl = file.getUrl();
  
  Logger.log("Saved map URL: " + imageUrl);

  var recipient = "andreysebastian10.g@gmail.com"; // Replace with recipient's email
  var subject = "Your Minimalist Route Map";
  var body = "<p>Here is your route:</p><img src='" + imageUrl + "' width='600'>";

  MailApp.sendEmail(
    {
    to: recipient,
    subject: subject,
    htmlBody: body
  });
}

