/*
Copyright 2024 Charles Villegas (for McGill Students Running Club)

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

/** VERIFY CONSTANTS AND UPDATE (IF APPLICABLE) */
const POINTS_EMAIL_NAME = 'Points Email Template';

const EMAIL_LEDGER_TARGETS = {
  feeStatus : LEDGER_INDEX.FEE_STATUS,
  firstName : LEDGER_INDEX.FIRST_NAME,
  totalPoints : LEDGER_INDEX.TOTAL_POINTS,
  lastRunDate : LEDGER_INDEX.LAST_RUN_DATE,
  runStreak : LEDGER_INDEX.RUN_STREAK,
  totalRuns : LEDGER_INDEX.TOTAL_RUNS,
  totalDistance : LEDGER_INDEX.TOTAL_DISTANCE,
  // TOTAL_ELEVATION,
};


function fTest() {
  const items = ["apple", "banana", "cherry"];

  function logItemWithPrefix(prefix, item, index) {
    console.log(`${prefix} Index ${index}: ${item}`);
  }

  // Using .bind() to pre-set the first argument
  items.forEach(logItemWithPrefix.bind(null, "Fruit"));
}

// Return latest log values
function getLatestLog() {
  return getLogInRow();
}

function getLogInRow(row = getValidLastRow(LOG_SHEET)) {
  const sheet = LOG_SHEET;
  const numCols = sheet.getLastColumn();
  return sheet.getSheetValues(row, 1, 1, numCols)[0];
}

function getLogAttendees(row) {
  // Get log attendees using stored index
  const attendeesCol = LOG_INDEX.ATTENDEE_NAME_EMAIL - 1;
  const thisLog = getLogInRow(row);

  // Return log attendees
  return thisLog[attendeesCol];

  /** Ensure input is not falsy and does not contain "None" */
  function attendeeFilter(input) {
    return input && !/\bNone\b/i.test(input);
  }
}

function getLedgerData() {
  const pointSheet = LEDGER_SHEET;
  
  // Define dimensions of sheet data
  const startCol = 1;
  const startRow = 2;
  const numRows = getValidLastRow(pointSheet) - 1;   // Remove header row
  const numCols = LEDGER_COL_COUNT;   // Exclude event-specific points

  const data = pointSheet.getSheetValues(startRow, startCol, numRows, numCols)
  console.log(data);
  return data;
}

function logStatus(messageArr, logSheet = LOG_SHEET, thisRow = getValidLastRow(logSheet)) {
  // Update the status of sending email
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss] ---');
  const statusRange = logSheet.getRange(thisRow, LOG_INDEX.EMAIL_STATUS);

  // Append status to previous value (if non-empty)
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  const updatedStatus = `${previousValue}${currentTime}\n${messageArr.join('\n')}`
  statusRange.setValue(updatedStatus);
}


/**
 * Function to send email to each member updating them on their points
 *
 * @trigger New headrun submission  // OLD: The 1st and 14th of every month
 * 
 * @todo only send email to attendee after every head run
 * @add stats: distance, elevation, duration, map, average_speed (pace), max_speed + points
 * 
 * @todo remove references to external sheets, i.e. attendance sheet
 * 
 * @todo Finish function migration
 *
 * @author [Charles Villegas](<charles.villegas@mail.mcgill.ca>) & ChatGPT
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 5, 2024
 * @update  Mar 23, 2025
 */

function pointsEmail() {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error ('Please switch to the McRUN Google Account before sending emails');
  }

  const logSheet = LOG_SHEET;
  const row = getValidLastRow(logSheet);

  // Get attendees from log
  const attendees = getLogAttendees(row);

  if (attendees.length === 0) {
    // Save return status of function execution
    return logMessages([`No recipients found for row: ${row}`], logSheet, row);
  }

  // Store recipient information as `{email : name}`
  // @todo Collect emails only in arr?
  const recipientMap = {};
  attendees.split('\n').forEach(entry => {
    const [name, email] = entry.split(':');
    recipientMap[email] = name;
  });

  // Get all names and point values from points, and names and emails from emails
  // Leave ledgerData as Array instead of Object for optimization
  const ledgerData = getLedgerData();
  const returnStatus = [];

  // Loop through emails, package member data, then send email
  for (const email of Object.keys(recipientMap)) {
    const entry = getLedgerEntry(email, ledgerData);
    const memberStats = extractEmailValues_(entry);  // Get values for email
    returnStatus.push(sendStatsEmail_(email, memberStats));  // Save return status
  }

  // Save return status of previous function execution
  logMessages(returnStatus, logSheet, row);
  return Logger.log(`Successfully executed 'pointsEmail' and logged messages in sheet`);

  /** Helper function */
  function extractEmailValues_(entry) {
    return Object.fromEntries(
      Object.entries(EMAIL_LEDGER_TARGETS).map(
        ([label, index]) => [label, entry[index - 1]]) // Convert 1-based index to 0-based
    );
  }
}

function sendStatsEmail_(email, memberStats, stravaActivity = {}) {
  // Prepare the HTML body from the template
  const template = HtmlService.createTemplateFromFile(POINTS_EMAIL_NAME);

  // Ensure placeholder names are valid
  template.FIRST_NAME = memberStats['firstName'];
  template.FEE_STATUS = memberStats['feeStatus'];
  template.MEMBER_POINTS = memberStats['totalPoints'];
  template.RUN_STREAK = memberStats['runStreak'];
  template.TOTAL_RUNS = memberStats['totalRuns'];
  template.TOTAL_DISTANCE = memberStats['totalDistance'];

  // Returns string content from constructed html template
  const pointsEmailHTML = template.evaluate().getContent();

  const subject = `Your Member Stats`;
  MailApp.sendEmail({
    //to: email,
    to: 'andrey.gonzalez@mail.mcgill.ca',
    subject: subject,
    htmlBody: pointsEmailHTML
  });

  // Log confirmation for the sent email with values for each variable
  return `Email sent to ${email} with ${memberStats['totalPoints']} points.`;
}


function mailMemberPoints_(trimmedName, email, points) {
  // Exit if no email found for member
  if (!email) {
    return Logger.log(`No email found for ${trimmedName}.`);
  }

  // Prepare the HTML body from the template
  const template = HtmlService.createTemplateFromFile(POINTS_EMAIL_NAME);
  template.FIRST_NAME = firstName;
  template.MEMBER_POINTS = points;

  // Returns string content from populated html template
  const pointsEmailHTML = template.evaluate().getContent();

  // Construct and send the email
  const subject = `Your Points Update`;

  MailApp.sendEmail({
    //to: email,
    to: 'andrey.gonzalez@mail.mcgill.ca',
    subject: subject,
    htmlBody: pointsEmailHTML
  });

  // Log confirmation for the sent email with values for each variable
  Logger.log(`Email sent to ${trimmedName} at ${email} with ${points} points.`);
}
