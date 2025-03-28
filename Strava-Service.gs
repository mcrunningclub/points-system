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
 * @author [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 23, 2025
 */

function safeReset() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const key = 'IS_RESET_ALLOWED';
  const isAllowed = scriptProperties.getProperty(key);

  if (!!isAllowed) {
    return Logger.log(`Please set '${key}' to true before trying again.`); // Log and exit
  }

  reset_();
  scriptProperties.setProperty(key, false);
  Logger.log(`Successfully reset service. Value of ${key}: false`);
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

