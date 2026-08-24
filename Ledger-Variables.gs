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
const LEDGER_COL = {
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
const LEDGER_COL_COUNT = Object.keys(LEDGER_COL).length;


/**
 * Maps columns to column number in log sheet
 */
const LOG_COL = {
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
 * Maps Strava stats to their formatting functions
 */
const NUMBER_FORMAT_MAP = {
  'distance': x => toFixedTruncate_(x, 2),
  'moving_time': x => toMinuteSeconds_(x),
  'average_speed': x => toMinuteSeconds_(x),
  'max_speed': x => x.toFixed(1),
  'total_elevation_gain': x => {
    const sign = (x > 0) ? '+' : '';
    return `${sign}${x.toFixed(0)}`;
  }
}


/**
 * Maps Strava stats to unit conversion factors for metric and imperial system
 * 
 * Distance -> convert meters to km or mile. 
 * Moving time -> keep the same.
 * Average speed -> convert meters/sec to km/sec or mi/sec. 
 * Max speed -> convert meters/sec to km/h or mph. 
 * Total elevation gain -> convert meters to feet for imperial.
 */
const UNITS_MAP = {
  'distance': { metric: 0.001, imperial: 1 / 1609.344 },
  'moving_time': { metric: 1, imperial: 1 },
  'average_speed': { metric: 0.001, imperial: 1 / 1609.344 },
  'max_speed': { metric: 3.6, imperial: 2.237 },
  'total_elevation_gain': { metric: 1, imperial: 3.2808 },
}


/**
 * Maps Strava stats to their target column in the event log sheet.
 */
const LOG_TARGETS = {
  'id': LOG_COL.STRAVA_ACTIVITY_ID,      // (long)
  'name': LOG_COL.STRAVA_ACTIVITY_NAME,  // (string)
  'distance': LOG_COL.DISTANCE_STRAVA,   // meters (float)
  'moving_time': LOG_COL.MOVING_TIME,  // seconds (int)
  'average_speed': LOG_COL.PACE,         // m per sec (float)
  'max_speed': LOG_COL.MAX_SPEED,        // m per sec (float)
  'total_elevation_gain': LOG_COL.ELEVATION,   // meters (float)
  'map': LOG_COL.MAP_POLYLINE,
  'mapUrl': LOG_COL.MAP_URL,
};

/**
 * Base url for the Strava API
 */
const STRAVA_BASE_URL = 'https://www.strava.com/api/v3/';

/**
 * Endpoint for Strava activities for the Strava API
 */
const ACTIVITIES_ENDPOINT = 'athlete/activities';

/**
 * Google Drive folder to store maps in
 */
const MAPS_FOLDER = 'run_maps';

/**
 * Base URL for Google Maps API
 */
const MAPS_BASE_URL = "https://maps.googleapis.com/maps/api/staticmap";

/**
 * Base URL of the Google Cloud Storage API
 */
const BASE_UPLOAD_URL = "https://storage.googleapis.com/upload/storage/v1/b";

/**
 * Name of bucket in Google Cloud Storage
 */
const STORAGE_BUCKET_NAME = 'run-map-storage.firebasestorage.app';