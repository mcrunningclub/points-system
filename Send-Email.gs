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
const MAPS_BASE_URL = "https://maps.googleapis.com/maps/api/staticmap?";

const EMAIL_LEDGER_TARGETS = {
  feeStatus : LEDGER_INDEX.FEE_STATUS,
  firstName : LEDGER_INDEX.FIRST_NAME,
  totalPoints : LEDGER_INDEX.TOTAL_POINTS,
  lastRunDate : LEDGER_INDEX.LAST_RUN_DATE,
  runStreak : LEDGER_INDEX.RUN_STREAK,
  totalRuns : LEDGER_INDEX.TOTAL_RUNS,
  totalDistance : LEDGER_INDEX.TOTAL_DISTANCE,
  totalElevation: LEDGER_INDEX.TOTAL_ELEVATION,
};


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

function sendTestEmail() {
  const template = HtmlService.createTemplateFromFile('Stats Email Template');

  // Returns string content from populated html template
  const templateHTML = template.evaluate().getContent();

  // Construct and send the email
  const subject = `Your Run Recap TEST`;

  MailApp.sendEmail({
    to: 'andreysebastian10.g@gmail.com',
    cc: 'andrey.gonzalez@mail.mcgill.ca',
    subject: subject,
    htmlBody: templateHTML
  });
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


/** Actual function that sends email */
// ⭐️⭐️⭐️⭐️

function sendStatsEmail_(email, memberStats) {
  const emailTemplate = STATS_EMAIL_OBJ;  // String instead of HTML template

  try {
    const msgObj = fillInTemplateFromObject_(emailTemplate, data);

    MailApp.sendEmail(
      //to: email
      'andreysebastian10.g@gmail.com',
      emailTemplate.subject,
      msgObj.text, { 
        htmlBody: msgObj.html,
      }
    );

  } catch(e) {
    return Logger.log(e);
  }

  // Log confirmation for the sent email with values for each variable
  return `Email sent to ${email} with ${memberStats['totalPoints']} points.`;
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
 * @update  Mar 27, 2025
 */

function sendStatsEmail(logSheet = LOG_SHEET, row = getValidLastRow(logSheet)) {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error ('Please switch to the McRUN Google Account before sending emails');
  }

  // Get attendees from log
  const attendees = getLogAttendees_(row);

  if (attendees.length === 0) {
    // Save return status of function execution
    return logMessages([`No recipients found for row: ${row}`], logSheet, row);
  }

  // Extract email and store in arr
  const recipientArr = 
    attendees.split('\n').reduce((acc, entry) => {
      const [, email] = entry.split(':');
      acc.push(email);
      return acc;
    }, []
  );

  // Get all names and point values from points, and names and emails from emails
  // Leave ledgerData as Array instead of Object for optimization
  const ledgerData = getLedgerData_();
  const returnStatus = [];

  // Loop through emails, package member data, then send email
  for (const email of recipientArr) {
    const entry = getLedgerEntry(email, ledgerData);
    const memberStats = extractEmailValues_(entry);  // Get values for post-run email
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

  // points spreadsheet (currently the test page)
  const POINTS_SHEET = LEDGER_SS.getSheetByName("test2");
  // columns (0 indexed)
  const EMAIL_COL = 0;
  const FNAME_COL = 2;
  const LAST_RUN_COL = 8;

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
function sendReminderEmail_(name, email) {
  // set up email using template
  const template = HtmlService.createTemplateFromFile('reminderemail');
  template.FIRST_NAME = name;
  let filledTemplate = template.evaluate();

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


/**
 * Google Maps API for headrun map creation
 */

function buildPostUrl_(polyline, imgSize = "580x420") {
  // URL Parameters
  const propertyStore = PropertiesService.getScriptProperties();
  const apiKey = propertyStore.getProperty(SCRIPT_PROPERTY_KEYS.googleMapAPI); // Replace with your API Key

  const googleCloudMapId = 'bfeadd271a2b0a58';  //'2ff6c54f4dd84b16';

  //Google Static Maps API URL
  return MAPS_BASE_URL +
    `size=${imgSize}&` +
    `map_id=${googleCloudMapId}&` +
    `path=enc:${polyline}&` + 
    `key=${apiKey}`;
}


function postToMakeWebhook_(postUrl, timestamp) {
  const webhookUrl = "https://hook.us1.make.com/8obb3hb6bzwgi7s4nyi8yfghb3kxsksc";
  const payload = JSON.stringify({ url: postUrl, name : timestamp });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  Logger.log("Response: " + response.getContentText());
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

Helpher function to fill email template.
*/

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{{}}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{{}}]+/g, "")] || "");
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

