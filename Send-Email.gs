/*
Copyright 2024 Charles Villegas (for McGill Students Running Club)

Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Copyright 2025 Mona Liu (for McGill Students Running Club)

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
const POINTS_EMAIL_SUBJECT = "Here's your post-run report! 🙌";
const EMAIL_SENDER_NAME = "McGill Students Running Club";

const EMAIL_LEDGER_TARGETS = {
  'FIRST_NAME': LEDGER_INDEX.FIRST_NAME,
  'USE_METRIC': LEDGER_INDEX.USE_METRIC,
  'TPOINTS': LEDGER_INDEX.TOTAL_POINTS,
  'LAST_RUN_DATE': LEDGER_INDEX.LAST_RUN_DATE,
  'TWEEKS': LEDGER_INDEX.RUN_STREAK,
  'TRUNS': LEDGER_INDEX.TOTAL_RUNS,
  'TOTAL_DISTANCE': LEDGER_INDEX.TOTAL_DISTANCE,
  'TOTAL_ELEVATION': LEDGER_INDEX.TOTAL_ELEVATION,
};

const EMAIL_PLACEHOLDER_LABELS = {
  'distance': 'DISTANCE',
  'elapsed_time': 'DURATION',
  'average_speed': 'PACE',
  'total_elevation_gain': 'ELEVATION',
  'max_speed': 'MSPEED',
  'mapUrl': 'MAP_URL',
  'id': 'ACTIVITY_ID',
  'points': 'POINTS',
  'mapCid': 'MAP_CID',
  'mapBlob': 'MAP_BLOB',
}


/** 
 * Testing Runtime Functions
 */

function fTest() {
  const items = ["apple", "banana", "cherry"];

  // Using .bind() to pre-set the first argument
  items.forEach(logItemWithPrefix.bind(null, "Fruit"));

  function logItemWithPrefix(prefix, item, index) {
    console.log(`${prefix} Index ${index}: ${item}`);
  }
}


function logStatus_(messageArr, logSheet = LOG_SHEET, thisRow = getValidLastRow(logSheet)) {
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
 * @trigger  New headrun submission  // OLD: The 1st and 14th of every month
 *
 * @author [Charles Villegas](<charles.villegas@mail.mcgill.ca>) & ChatGPT
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 5, 2024
 * @update  Apr 3, 2025
 */

function sendStatsEmail(logSheet = LOG_SHEET, row = getValidLastRow(logSheet)) {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error('Please switch to the McRUN Google Account before sending emails');
  }

  // Get attendees from log
  const attendees = getAttendeesInLog_(row);
  if (!attendees) {
    Logger.log(`No recipients found for row: ${row}`);
    return null;
  }

  // Get activity and add headrun points from log
  const activityStats = findAndStoreStravaActivity(row);
  activityStats['points'] = getEventPointsInRow_(row);

  // Extract email and store in arr
  const recipientArr =
    attendees.split('\n').reduce((acc, entry) => {
      const [, email] = entry.split(':');
      acc.push(email);
      return acc;
    }, []
  );

  const returnStatus = emailMemberStats_(recipientArr, activityStats);

  // Print log and save return status of `emailMemberStats`
  console.log(activityStats);
  logStatus_(returnStatus, logSheet, row);
  Logger.log(`Successfully executed 'sendStatsEmail' and logged messages in sheet`);
}


/** Helper 1: Send member stats to recipient */
function emailMemberStats_(recipients, activity) {
  // Get all names and point values from points, and names and emails from emails
  // Leave ledgerData as Array instead of Object for optimization
  const ledgerData = GET_LEDGER_();
  const isEmailAllowed = LEDGER_INDEX.EMAIL_ALLOWED--;    // Make 0-indexed for arr
  const res = [];

  // Get activity stats in metric and US imperial
  const allStats = convertAndFormatStats_(activity);

  // Transform key labels in Strava to placeholder names in email
  const { metric: metricStats, imperial: imperialStats } = filterEmailValues(allStats);

  // Loop through emails, package member data, then send email
  for (const email of recipients) {
    const entry = getLedgerEntry(email, ledgerData);

    if (!entry[isEmailAllowed]) continue;   // Only sent to members who consented

    const memberTotalStats = sheetToEmailLabels(entry);  // Get values for post-run email
    const preferredStats = memberTotalStats['USE_METRIC'] ? metricStats : imperialStats;

    // Email report and log response
    res.push(emailReport_(email, { ...memberTotalStats, ...preferredStats }));
  }

  return res;

  /** Helper: Package run stats using ledger and `EMAIL_LEDGER_TARGETS` */
  function sheetToEmailLabels(entry) {
    return Object.fromEntries(
      Object.entries(EMAIL_LEDGER_TARGETS).map(
        ([label, index]) => [label, entry[index - 1]]) // Convert 1-based index to 0-based
    );
  }

  function filterEmailValues(data) {
    const ret = { metric: {}, imperial: {} };
    const systems = Object.keys(ret);

    for (const [objKey, emailKey] of Object.entries(EMAIL_PLACEHOLDER_LABELS)) {
      systems.forEach(sys => {
        ret[sys][emailKey] = data[sys][objKey] || "";
      });
    }

    return ret;
  }
}


function emailReport_(email, memberStats) {
  // Create template to populate
  const template = HtmlService.createTemplateFromFile('Points Email V2');

  // Get member's system preference to format email
  const useMetric = memberStats['USE_METRIC'];
  template.USE_METRIC = memberStats['USE_METRIC'];

  // Populate member's general stats
  template.FIRST_NAME = memberStats['FIRST_NAME'];
  template.TPOINTS = memberStats['TPOINTS'];
  template.TWEEKS = memberStats['TWEEKS'];
  template.TRUNS = memberStats['TRUNS']

  // Populate activity units
  template.DISTANCE = memberStats['DISTANCE'];
  template.DURATION = memberStats['DURATION'];
  template.PACE = memberStats['PACE']
  template.ELEVATION = memberStats['ELEVATION'];
  template.MSPEED = memberStats['MSPEED'];
  template.POINTS = memberStats['POINTS'];
  template.ACTIVITY_ID = memberStats['ACTIVITY_ID'];
  template.MAP_URL = memberStats['MAP_URL'];

  // Evaluate template and log message
  const filledTemplate = template.evaluate();
  Logger.log(`Now constructing email with ${useMetric ? 'metric' : 'imperial'} units.`);

  MailApp.sendEmail(
    message = {
      to: email,
      bcc: 'andrey.gonzalez@mail.mcgill.ca',
      name: EMAIL_SENDER_NAME,
      subject: POINTS_EMAIL_SUBJECT,
      replyTo: MCRUN_EMAIL,
      htmlBody: filledTemplate.getContent(),
    }
  );

  // Log confirmation for the sent email with member stats
  const log = `Stats email sent to ${email}`;
  Logger.log(log);
  return log;
}


/**
 * Automatically triggered to send reminder email to members whose
 * "last run" date is over 2 weeks ago
 * 
 * @trigger every Monday
 * 
 * @author Mona Liu <mona.liu@mail.mcgill.ca>
 * 
 * @date 2025/03/30
 */
function checkAndSendWinBackEmail() {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error('Please switch to the McRUN Google Account before sending emails');
  }

  // points spreadsheet (currently the test page)
  const POINTS_SHEET = LEDGER_SS.getSheetByName("test2");
  // columns (0 indexed)
  const EMAIL_COL = LEDGER_INDEX.EMAIL--;    // const EMAIL_COL = 0;
  const FNAME_COL = LEDGER_INDEX.FIRST_NAME--;   // const FNAME_COL = 2;
  const LAST_RUN_COL = LEDGER_INDEX.LAST_RUN_DATE;   // const LAST_RUN_COL = 8;

  // ^-- I refactored your code to use the centralized indices of the sheet
  // to future-proof any columns adds/removes


  // make date object for 2 weeks ago
  let dateThreshold = new Date();
  dateThreshold.setDate(dateThreshold.getDate() - 14);

  // get all data entries as 2d array (row, col)
  let allMembers = POINTS_SHEET.getDataRange().getValues();

  // loop through member entries (questionable efficiency)
  // except first row which is the header
  for (let i = 1; i < allMembers.length; i++) {
    // check for last run date
    let member = allMembers[i];
    let lastRunAsStr = member[LAST_RUN_COL];

    // skip rows with no data
    if (lastRunAsStr != '') {
      // convert last run date into date object
      let lastRunAsDate = new Date(lastRunAsStr);

      // send reminder email if needed
      if (lastRunAsDate < dateThreshold) {
        sendReminderEmail_(member[FNAME_COL], member[EMAIL_COL]);
      }
    }
  }
}


/**
 * Creates reminder email from member name and template,
 * sends it to given address
 * 
 * @param {String} name Member's first name
 * @param {String} email Member's email address
 * @returns None
 * 
 * @author Mona Liu <mona.liu@mail.mcgill.ca>
 * 
 * @date 2025/03/30
 */
function sendWinBackEmail_(name, email) {
  // set up email using template
  const template = HtmlService.createTemplateFromFile('winbackemail');
  template.FIRST_NAME = name;
  let filledTemplate = template.evaluate();

  // send email
  try {
    MailApp.sendEmail(
      message = {
        to: email,
        name: EMAIL_SENDER_NAME,
        subject: "We've missed you!",
        htmlBody: filledTemplate.getContent()
      }
    );

  } catch (e) {
    Logger.log(e);
  }

  // Log confirmation for the sent email
  Logger.log(`Win-back email sent to ${email}.`);
}


/** 
 * Function to send first iteration of Stats Email Template.
 * 
 * @author [Charles Villegas](<charles.villegas@mail.mcgill.ca>) & ChatGPT
 * @deprecated
 */

function mailMemberPointsV1_(trimmedName, email, points) {
  // Exit if no email found for member
  if (!email) {
    return Logger.log(`No email found for ${trimmedName}.`);
  }

  // Prepare the HTML body from the template
  const template = HtmlService.createTemplateFromFile('Points Email V1');
  template.FIRST_NAME = firstName;
  template.MEMBER_POINTS = points;

  // Returns string content from populated html template
  const pointsEmailHTML = template.evaluate().getContent();

  // Construct and send the email
  const subject = `Your Points Update`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: pointsEmailHTML,
    name: EMAIL_SENDER_NAME,
  });

  // Log confirmation for the sent email with values for each variable
  Logger.log(`Email sent to ${trimmedName} at ${email} with ${points} points.`);
}


/*
Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Helper function to fill email template.

- Added explicit string conversion of values for `escapeData`.
*/

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 * 
 * @update  Explicit string conversion of values for `escapeData`.
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{{}}]+}}/g, key => {
    return escapeData_(`${data[key.replace(/[{{}}]+/g, "")]}` || "");
  });


  return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};

