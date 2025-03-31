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
const POINTS_EMAIL_NAME = 'Stats Email Template';
const MAPS_BASE_URL = "https://maps.googleapis.com/maps/api/staticmap";

const EMAIL_LEDGER_TARGETS = {
  'FIRST_NAME' : LEDGER_INDEX.FIRST_NAME,
  'TPOINTS' : LEDGER_INDEX.TOTAL_POINTS,
  'LAST_RUN_DATE' : LEDGER_INDEX.LAST_RUN_DATE,
  'TWEEKS' : LEDGER_INDEX.RUN_STREAK,
  'TRUNS' : LEDGER_INDEX.TOTAL_RUNS,
  'TOTAL_DISTANCE' : LEDGER_INDEX.TOTAL_DISTANCE,
  'TOTAL_ELEVATION' : LEDGER_INDEX.TOTAL_ELEVATION,
};

const EMAIL_PLACEHOLDER_LABELS = {
  'distance' : 'DISTANCE',
  'elapsed_time' : 'DURATION',
  'average_speed' : 'PACE',
  'total_elevation_gain' : 'ELEVATION',
  'max_speed' : 'MSPEED',
  'mapUrl' : 'RUN_MAP',
  'id' : 'ACTIVITY_ID',
  'points' : 'POINTS'
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
 * @author2 [Andrey S Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 5, 2024
 * @update  Mar 31, 2025
 */

function sendStatsEmail(logSheet = LOG_SHEET, row = getValidLastRow(logSheet)) {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error ('Please switch to the McRUN Google Account before sending emails');
  }

  // Get attendees from log
  const attendees = getLogAttendees_(row);
  const activityStats = findAndStoreStravaActivity(row);

  // Add headrun points
  activityStats['points'] = 50;   //todo: REMOVE HARD-CODING

  console.log(activityStats);

  if (!attendees) {
    return null && Logger.log(`No recipients found for row: ${row}`);
  }

  // Extract email and store in arr
  const recipientArr = 
    attendees.split('\n').reduce((acc, entry) => {
      const [, email] = entry.split(':');
      acc.push(email);
      return acc;
    }, []
  );

  const returnStatus = emailMemberStats_(recipientArr, activityStats);

  // Save return status of previous function execution
  logStatus_(returnStatus, logSheet, row);
  Logger.log(`Successfully executed 'pointsEmail' and logged messages in sheet`);
}


/** Helper 1: Send member stats to recipient */
function emailMemberStats_(recipients, activity) {
  // Get all names and point values from points, and names and emails from emails
  // Leave ledgerData as Array instead of Object for optimization
  const ledgerData = GET_LEDGER_();
  const res = [];

  // Transform key labels in Strava to placeholder names in email
  convertAllUnits_(activity, true);
  const targetStats = prepareEmailFields(activity);

  // Loop through emails, package member data, then send email
  for (const email of recipients) {
    const entry = getLedgerEntry(email, ledgerData);
    const memberTotalStats = sheetToEmailLabels(entry);  // Get values for post-run email
    res.push(emailReport_(email, {...memberTotalStats, ...targetStats}));
  }
  return res;

  /** Helper: Package run stats using ledger and `EMAIL_LEDGER_TARGETS` */
  function sheetToEmailLabels(entry) {
    return Object.fromEntries(
      Object.entries(EMAIL_LEDGER_TARGETS).map(
        ([label, index]) => [label, entry[index - 1]]) // Convert 1-based index to 0-based
    );
  }

  function prepareEmailFields(data) {
    return Object.entries(EMAIL_PLACEHOLDER_LABELS).reduce((acc, [objKey, emailKey]) => {
      acc[emailKey] = data[objKey] || "";
      return acc;
    }, {});
  }
}


/** ⭐️ Actual function that sends email ⭐️ */
function emailReport_(email, memberStats) {
  const emailTemplate = STATS_EMAIL_OBJ;  // String instead of HTML template

  // Append general data to activity stats (e.g. current year)
  const generalData = {'THIS_YEAR' : `${new Date().getFullYear()}`};
  memberStats = {...memberStats, ...generalData};

  const msgObj = fillInTemplateFromObject_(emailTemplate, memberStats);

  MailApp.sendEmail({
    //to: email
    to : 'andrey.gonzalez@mail.mcgill.ca',
    subject: emailTemplate.subject,
    htmlBody: msgObj.html,
    name: 'McGill Students Running Club'
  });

  // Log confirmation for the sent email with member stats
  return `Stats email sent to ${email}.`;
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
function checkAndSendReminderEmail() {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error('Please switch to the McRUN Google Account before sending emails');
  }

  // get current date

  // check all member entries who have a "last run" date
  // make list of emails and first names?

  // loop through members and get info


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
function sendReminderEmail_(name, email) {
  // set up email using template
  const template = HtmlService.createTemplateFromFile('reminderemail');
  template.FIRST_NAME = name;
  const filledTemplate = template.evaluate();

  // send email
  try {
    MailApp.sendEmail(
      message = {
        to: email,
        name: "McRUN",
        subject: "We've missed you!",
        htmlBody: filledTemplate.getContent()
      }
    );

  } catch (e) {
    Logger.log(e);
  }

  // Log confirmation for the sent email
  Logger.log(`Reminder email sent to ${email}.`);
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
  const template = HtmlService.createTemplateFromFile(POINTS_EMAIL_NAME);
  template.FIRST_NAME = firstName;
  template.MEMBER_POINTS = points;

  // Returns string content from populated html template
  const pointsEmailHTML = template.evaluate().getContent();

  // Construct and send the email
  const subject = `Your Points Update`;

  MailApp.sendEmail({
    //to: email,
    subject: subject,
    htmlBody: pointsEmailHTML
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

