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
const LEDGER_SHEET_NAME = 'Member Points';
const LEDGER_SHEET = LEDGER_SS.getSheetByName(LEDGER_SHEET_NAME);

const LOG_SHEET_NAME = 'Event Log';
const LOG_SHEET = LEDGER_SS.getSheetByName(LOG_SHEET_NAME);

const TIMEZONE = getUserTimeZone_();
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';

/** RUN LEVELS + COUNT */
const ATTENDEE_MAP = {
  // 'beginner': ATTENDEES_BEGINNER_COL,
  // //'easy': ATTENDEES_BEGINNER_COL,
  // 'intermediate': ATTENDEES_INTERMEDIATE_COL,
  // 'advanced':  ATTENDEES_ADVANCED_COL,
};

const LEVEL_COUNT = Object.keys(ATTENDEE_MAP).length;

/** EXTERNAL SHEETS */
const ATTENDANCE_SHEET_NAME = 'HR Attendance W25';
const ATTENDANCE_URL = 'https://docs.google.com/spreadsheets/d/1SnaD9UO4idXXb07X8EakiItOOORw5UuEOg0dX_an3T4/';


/** STORES INDEX OF COLUMNS IN POINTS_SHEET */
const LEDGER_INDEX = {
  EMAIL : 1,
  FEE_STATUS : 2,
  FIRST_NAME : 3,
  LAST_NAME : 4,
  FULL_NAME : 5,
  TOTAL_POINTS : 6,
  REGISTERED_POINTS : 7,
  FEE_PAID_POINTS : 8,
  LAST_RUN_DATE : 9,
  RUN_STREAK : 10,
  TOTAL_RUNS : 11,
  TOTAL_DISTANCE : 12,
  TOTAL_ELEVATION : 13,
  // Cols 14+ are event-specific points
}

/** LEDGER SHEET COL SIZE (WITHOUT EVENT-SPECIFIC POINTS COL) */
const LEDGER_COL_COUNT = Object.keys(LEDGER_INDEX).length;


/** STORES INDEX OF COLUMNS IN LOG_SHEET */
const LOG_INDEX = {
  IMPORT_TIMESTAMP : 1,
  EVENT : 2,
  EVENT_TIMESTAMP : 3,
  ATTENDEE_NAME_EMAIL : 4,
  DISTANCE_ESTIMATED : 5,
  EVENT_POINTS : 6,
  EMAIL_STATUS : 7,
  STRAVA_ACCOUNT : 8,
  STRAVA_ACTIVITY_NAME : 9,
  DISTANCE_STRAVA : 10,
  ELAPSED_TIME : 11,
  PACE : 12,
  MAX_SPEED : 13,
  ELEVATION : 14,
  MAP_POLYLINE : 15,
}

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}

