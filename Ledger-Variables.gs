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
 * Ledger spreadsheet (entire file) object
 */
const LEDGER_SS = SpreadsheetApp.getActiveSpreadsheet();

/**
 * ID for the ledger spreadsheet (entire file)
 */
const LEDGER_SS_ID = '1sar-Pmfb_Nar0Lc9u8-rXyllLvQMqBFlSwolCoHX-_4';

/**
 * Name of the ledger sheet
 */
const LEDGER_SHEET_NAME = 'Member Points';

/**
 * Ledger sheet object
 */
const LEDGER_SHEET = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);

/**
 * Name for event log sheet
 */
const LOG_SHEET_NAME = 'Event Log';

/**
 * Event log sheet object
 */
const LOG_SHEET = LEDGER_SS.getSheetByName(LOG_SHEET_NAME);

/**
 * Ledger spreadsheet (entire file)
 */
let LEDGER_DATA = null;

/**
 * Gets contents of points ledger and stores it in the LEDGER_DATA constant
 */
const GET_LEDGER = () => {
  LEDGER_DATA = LEDGER_DATA ?? getLedgerData_();
  return LEDGER_DATA;
}

/**
 * Gets the log sheet by ID/name
 * ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
 * SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
 */
const GET_LOG_SHEET = () => {
  return (LOG_SHEET) ?? SpreadsheetApp.openById(LEDGER_SS_ID).getSheetByName(LOG_SHEET_NAME);
}

/**
 * Gets the ledger sheet by ID/name
 * ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
 * SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
 */
const GET_LEDGER_SHEET = () => {
  return (LEDGER_SHEET) ?? SpreadsheetApp.openById(LEDGER_SS_ID).getSheetByName(LEDGER_SHEET_NAME);
}

/** 
 * Timezone of the script
 * IMPORTANT FOR DATETIME FORMATTING AND SENDING EMAILS 
 */
const TIMEZONE = getUserTimeZone_();

/**
 * Official club email
 * IMPORTANT FOR DATETIME FORMATTING AND SENDING EMAILS
 */
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';


/** 
 * Keys of properties in script properties (MAKE SURE NAMES MATCHES ACTUAL STORE) 
 */
const SCRIPT_PROPERTY_KEYS = {
  clientID: 'CLIENT_ID',
  clientSecret: 'CLIENT_SECRET',
  googleMapAPI: 'GOOGLE_MAPS_API_KEY',
  googleCloudKey: 'GOOGLE_CLOUD_KEY',
  extraStrava : 'EXTRA_STRAVA_ARR',
  isSendAllowed : 'IS_SEND_ALLOWED',
  isResetAllowed: 'IS_RESET_ALLOWED,'
};

/** 
 * Maps columns to column number in points ledger sheet
 * Col 16+ store event-specific points
 */
const LEDGER_INDEX = {
  EMAIL: 1,
  FEE_STATUS: 2,
  FIRST_NAME: 3,
  LAST_REGISTERED: 4,
  FULL_NAME: 5,
  EMAIL_ALLOWED : 6,
  USE_METRIC : 7,
  TOTAL_POINTS: 8,
  REGISTRATION_POINTS: 9,
  FEE_PAID_POINTS: 10,
  LAST_RUN_DATE: 11,
  RUN_STREAK: 12,
  TOTAL_RUNS: 13,
  TOTAL_DISTANCE: 14,
  TOTAL_ELEVATION: 15
}

/**
 * LEDGER SHEET COL SIZE (WITHOUT EVENT-SPECIFIC POINTS COL)
 */
const LEDGER_COL_COUNT = Object.keys(LEDGER_INDEX).length;


/**
 * Maps columns to column number in log sheet
 */
const LOG_INDEX = {
  IMPORT_TIMESTAMP: 1,
  EVENT: 2,
  HEADRUNNERS: 3,
  EVENT_TIMESTAMP: 4,
  ATTENDEES: 5,
  DISTANCE_ESTIMATED: 6,
  EVENT_POINTS: 7,
  EMAIL_STATUS: 8,
  STRAVA_ACTIVITY_ID: 9,
  STRAVA_ACTIVITY_NAME: 10,
  DISTANCE_STRAVA: 11,
  MOVING_TIME: 12,
  PACE: 13,
  MAX_SPEED: 14,
  ELEVATION: 15,
  MAP_POLYLINE: 16,
  MAP_URL: 17,
}

/**
 * Gets the timezone of the script
 * 
 * @return {string} Time zone
 */
function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

/**
 * Gets the email of the user accessing the script
 * 
 * @return {string}  Email address
 */
function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}

/**
 * Tries to find a file in Google Drive by its name
 * 
 * May have an error if there are no results?
 * 
 * @param {string} name  Name of the file to find
 * @return {File}  The first result of the search
 */
function getFileByName_(name) {
  return DriveApp.searchFiles(`title contains '${name}'`).next();
}

/**
 * Tries to find a file in Google Drive by its ID
 * 
 * May have an error if there are no results?
 * 
 * @param {string} id  ID of the file to find
 * @return {File}  The first result of the search
 */
function getFileById_(id) {
  return DriveApp.getFileById(id);
}

/**
 * Creates log in the console with specific formatting
 * 
 * @param {string} msg  The message to log
 * @param {string} [funcName]  Name of the function returning the message, if applicable. Defaults to ""
 * @param {boolean} ][useLogger]  Whether to use the logger (true) or console.log (false). Defaults to true.
 */
function logAsPL_(msg, funcName = "", useLogger = true) {
  const message = `[PL#${funcName}] ${msg}`;
  useLogger ? Logger.log(message) : console.log(message);
}
