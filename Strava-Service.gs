/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Copyright 2016 Google Inc. All Rights Reserved. [Source](https://github.com/googleworkspace/apps-script-oauth2)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

THIS FILE HAS BEEN MODIFIED BY ANDREY GONZALEZ AS FOLLOWING:
- Moved confidential variables to project's script properties
- Added safety check to run `reset` using `safeReset`
- Renamed `getService_` to `getStravaService_`
- Created `callStravaAPI_` using original function `run`
- Modified variable names and declaration
- Improved documentation and inline comments
*/


/**
 * Reset the authorization state, so that it can be re-tested.
 */

function reset_() {
  var service = getStravaService_();
  service.reset();
}


/**
 * Run `reset` safely using script property flag `IS_RESET_ALLOWED`.
 * 
 * Must manually change value before running. Once allowed, flag toggles back to false.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 23, 2025
 */

function safeReset() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const key = 'IS_RESET_ALLOWED';
  const isAllowed = !!scriptProperties.getProperty(key);

  if (!isAllowed) {
    return Logger.log(`Please set '${key}' to true before trying again.`); // Log and exit
  }

  reset_();
  scriptProperties.setProperty(key, false);
  Logger.log(`Successfully reset service. Value of ${key}: false`);
}


/** Get Strava Activity in the input range */
function getStravaActivity_(fromTimestamp, toTimestamp) {
  // Package query for Strava API
  const queryObj = { 
    'after' : fromTimestamp, 
    'before' : toTimestamp,
    'include_all_efforts' : true 
  };

  const endpoint = ACTIVITIES_ENDPOINT;
  return callStravaAPI_(endpoint, queryObj);    // Returns a list of Objects
}


/**
 * Configures the service using the OAuth2 library.
 * 
 * Three required and optional parameters are not specified
 * because the library creates the authorization URL with them
 * automatically: `redirect_url`, `response_type`, and `state`.
 * 
 * #### APPENDED COMMENTS BY USER
 * 
 * Client ID and Secret stored in script properties. *(Mar 23, 2025)*
 * 
 * @see 'https://github.com/googleworkspace/apps-script-oauth2/blob/main/samples/Strava.gs'
 */

function getStravaService_() {
  // Get user-defined keys in Script Property for proper access
  const myScriptKeys = SCRIPT_PROPERTY_KEYS;

  // Save required script properties 
  const scriptProperties = PropertiesService.getScriptProperties();
  const clientId = scriptProperties.getProperty(myScriptKeys.clientID);
  const clientSecret = scriptProperties.getProperty(myScriptKeys.clientSecret);

  // Define scope of service to request (space-separated for Google services)
  const scope = 'activity:read_all,profile:read_all';

  // Create and return a new service called "Strava"
  return OAuth2.createService('Strava')

    /** Set the endpoint URL for Strava auth */
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')

    /** Set the client ID and secret */
    .setClientId(clientId)
    .setClientSecret(clientSecret)

    /** Set the name of the callback function `authCallback_` 
     * that should be invoked to complete the OAuth flow. */
    .setCallbackFunction('authCallback_')

    /** Set the property store where authorized tokens should be persisted */
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(scope)
    ;
}


/**
 * Handles the OAuth callback.
 * 
 * #### APPENDED COMMENTS BY USER
 * 
 * Must have global scope in project *(Mar 23, 2025)*
 * 
 */

function authCallback_(request) {
  var stravaService = getStravaService_();
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
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
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
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
  Logger.log('[PL] App has access.');

  // Get API endpoint
  endpoint = STRAVA_BASE_URL + endpoint;
  const query_string = queryObjToString_(query_object);

  const headers = {
    Authorization: 'Bearer ' + service.getAccessToken(),
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
 * @param {object} query_objec]  Param-value pair.
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
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 7, 2024
 * @update  Mar 31, 2025
 */

function queryObjToString_(query_object) {
  if (query_object.length === 0) return '';   // Check if object empty

  const query_string = Object.entries(query_object)
    .map(([param, value]) => `${param}=${value}`)
    .join('&');

  return '?' + query_string;
}


/** Functions previously in `Passkit-API.gs` */

function genTokenSign_(token, secret) {
  if (token.length != 2) {
      return;
  }
  var hash = Utilities.computeHmacSha256Signature(token.join("."), secret);
  var base64Hash = Utilities.base64Encode(hash);
  return urlConvertBase64_(base64Hash);
}

function base64url_(input) {
  var base64String = Utilities.base64Encode(input);
  return urlConvertBase64_(base64String);
}

function urlConvertBase64_(input) {
  var output = input.replace(/=+$/, '');
  output = output.replace(/\+/g, '-');
  output = output.replace(/\//g, '_');
  return output;
}

