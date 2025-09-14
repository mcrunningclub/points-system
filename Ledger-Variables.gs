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

/** SHEET CONSTANTS */
const LEDGER_SS = SpreadsheetApp.getActiveSpreadsheet();
const LEDGER_SS_ID = '1sar-Pmfb_Nar0Lc9u8-rXyllLvQMqBFlSwolCoHX-_4';

const LEDGER_SHEET_NAME = 'Member Points';
const LEDGER_SHEET = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);

const LOG_SHEET_NAME = 'Event Log';
const LOG_SHEET = LEDGER_SS.getSheetByName(LOG_SHEET_NAME);


let LEDGER_DATA = null;

const GET_LEDGER_ = () => {
  LEDGER_DATA = LEDGER_DATA ?? getLedgerData_();
  return LEDGER_DATA;
}

// ALLOWS PROPER SHEET REF WHEN ACCESSING AS LIBRARY FROM EXTERNAL SCRIPT
// SpreadsheetApp.getActiveSpreadsheet() DOES NOT WORK IN EXTERNAL SCRIPT
const GET_LOG_SHEET_ = () => {
  return (LOG_SHEET) ?? SpreadsheetApp.openById(LEDGER_SS_ID).getSheetByName(LOG_SHEET_NAME);
}

const GET_LEDGER_SHEET = () => {
  return (LEDGER_SHEET) ?? SpreadsheetApp.openById(LEDGER_SS_ID).getSheetByName(LEDGER_SHEET_NAME);
}

/** IMPORTANT FOR DATETIME FORMATTING AND SENDING EMAILS */
const TIMEZONE = getUserTimeZone_();
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';


/** SCRIPT PROPERTIES (MAKE SURE NAMES MATCHES ACTUAL STORE) */
const SCRIPT_PROPERTY_KEYS = {
  clientID: 'CLIENT_ID',
  clientSecret: 'CLIENT_SECRET',
  googleMapAPI: 'GOOGLE_MAPS_API_KEY',
  googleCloudKey: 'GOOGLE_CLOUD_KEY',
  extraStrava : 'EXTRA_STRAVA'
};


/** RUN LEVELS + COUNT */
const ATTENDEE_MAP = {
  // 'beginner': ATTENDEES_BEGINNER_COL,
  // //'easy': ATTENDEES_BEGINNER_COL,
  // 'intermediate': ATTENDEES_INTERMEDIATE_COL,
  // 'advanced':  ATTENDEES_ADVANCED_COL,
};

const LEVEL_COUNT = Object.keys(ATTENDEE_MAP).length;


/** STORES INDEX OF COLUMNS IN POINTS_SHEET */
const LEDGER_INDEX = {
  EMAIL: 1,
  FEE_STATUS: 2,
  FIRST_NAME: 3,
  LAST_NAME: 4,
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
  TOTAL_ELEVATION: 15,
  // Cols+ store event-specific points
}

/** LEDGER SHEET COL SIZE (WITHOUT EVENT-SPECIFIC POINTS COL) */
const LEDGER_COL_COUNT = Object.keys(LEDGER_INDEX).length;


/** STORES INDEX OF COLUMNS IN LOG_SHEET */
const LOG_INDEX = {
  IMPORT_TIMESTAMP: 1,
  EVENT: 2,
  EVENT_TIMESTAMP: 3,
  ATTENDEE_NAME_EMAIL: 4,
  DISTANCE_ESTIMATED: 5,
  EVENT_POINTS: 6,
  EMAIL_STATUS: 7,
  STRAVA_ACTIVITY_ID: 8,
  STRAVA_ACTIVITY_NAME: 9,
  DISTANCE_STRAVA: 10,
  MOVING_TIME: 11,
  PACE: 12,
  MAX_SPEED: 13,
  ELEVATION: 14,
  MAP_POLYLINE: 15,
  MAP_URL: 16,
}


function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}

function getFileByName_(name) {
  return DriveApp.searchFiles(`title contains '${name}'`).next();
}

function getFileById_(id) {
  return DriveApp.getFileById(id);
}

