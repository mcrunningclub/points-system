/*
Copyright 2024 Charles Villegas (for McGill Students Running Club)

Copyright 2024-25 Andrey Gonzalez (for McGill Students Running Club)

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

/**
 * Puts given messages (status of sending email) into the log sheet in the Email Status column
 * 
 * @param {string[]} messages  Array of messages to log
 * @param {SpreadsheetApp.Sheet} logSheet  Log sheet object
 * @param {integer} row  Row to log messages in
 */
function logStatus_(messages, logSheet, row) {
  // Update the status of sending email
  const currentTime = Utilities.formatDate(new Date(), TIMEZONE, '[dd-MMM HH:mm:ss] ---');
  const statusRange = logSheet.getRange(row, LOG_COL.EMAIL_STATUS);

  // Append status to previous value (if non-empty)
  const previousValue = statusRange.getValue() ? statusRange.getValue() + '\n' : '';
  const updatedStatus = `${previousValue}${currentTime}\n${messages.join('\n')}`
  statusRange.setValue(updatedStatus);
}

/**
 * Checks if email sending is allowed according to script properties and throws error if not.
 */
function isEmailSendingAllowed_() {
  const store = PropertiesService.getScriptProperties();
  const isSendAllowed = store.getProperty(SCRIPT_PROPERTY_KEYS.isSendAllowed);
  if (isSendAllowed === 'false') {
    throw new Error('[PL] Sending emails is currently toggled off. Please reset to continue.');
  }
}


/**
 * Function to send email to each member updating them on their points
 *
 * @trigger  New headrun submission  // OLD: The 1st and 14th of every month
 * 
 * @param {Spreadsheet.sheet} logSheet  Log sheet object
 * @param {integer} row  Row with the activity to send email for
 *
 * @author [Charles Villegas](<charles.villegas@mail.mcgill.ca>) & ChatGPT
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Nov 5, 2024
 * @update  Sep 25, 2025
 */

function checkAndSendPostRunEmail(logSheet = GET_LOG_SHEET(), row = getValidLastRow_(logSheet)) {
  const funcName = checkAndSendPostRunEmail.name;
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != MCRUN_EMAIL) {
    throw new Error(`[PL#${checkAndSendPostRunEmail.name}] Please switch to the McRUN Google Account before sending emails`);
  }

  isEmailSendingAllowed_();    // throws error if not allowed
  logAsPL_(`Email can be sent! Continuing execution now...`, funcName);

  // Get attendees from log
  const attendees = getAttendeesInRow_(row);
  if (!attendees) {
    logAsPL_(`No recipients found for row: ${row}`, funcName);
    return null;
  }

  // Get activity and add headrun points from log
  const activityStats = findAndStoreStravaActivity(row);
  if (!activityStats) return;   // Cannot send email without stats
  logAsPL_(`Found activity stats!`, funcName);

  // Otherwise send email with extracted stats
  activityStats['points'] = getEventPointsInRow_(row);

  // Append headrun details
  const { date, level } = getDateAndLevelInRow_(row);
  activityStats['run_date'] = date;
  activityStats['level'] = level;

  // Add headrunners
  const headrunnerInfo = splitInfo(getHeadrunnersInRow_(row));
  activityStats['headrunners'] = headrunnerInfo?.names.join(', ').replace(/,([^,]*)$/, ' & $1');

  // Extract email and store in arr
  const recipientArr = splitInfo(attendees)?.emails;
  const copyRecipientArr = headrunnerInfo?.emails;

  logAsPL_(`Now trying to execute '${sendPostRunEmailsForActivity_.name}'...`, funcName);
  const returnStatus = sendPostRunEmailsForActivity_(recipientArr, activityStats);

  // Print log and save return status of `emailMemberStats`
  console.log(activityStats);
  logStatus_(returnStatus, logSheet, row);
  logAsPL_("Successfully executed and logged messages in sheet", funcName);

  /** Helper to get names and emails formatted as `[Bob Burger|Bob B.]:bob@mail.com`*/
   function splitInfo(infoAsStr) {
    const splitByNewline = infoAsStr.split('\n');

    // First try spliting
    try{
      const res = splitByNewline.reduce((acc, entry) => {
        const [name, email=""] = entry.split(':');
        acc.names.push(name.trim());
        acc.emails.push(email.trim());
        return acc;
      }, {names: [], emails: []});
      
      return res || {};
    }
    catch(e) {
      logAsPL_(`Could not split headrunners by ':'; returning names only`);
      logAsPL_(`Catch error: ${e.message}`);
      return { names: splitByNewline, emails: [] };
    }
  }
}


/** 
 * Sends post run email to all recipients for the specified activity
 * 
 * @param {string[]} recipients  Email addresses to send email to
 * @param {Object} activity  Activity stats
 * 
 * @return {string[]}  List of status messages indicating whether each email was sent successfully
 */
function sendPostRunEmailsForActivity_(recipients, activity) {
  // Get all names and point values from points, and names and emails from emails
  // Leave ledgerData as Array instead of Object for optimization
  const ledgerData = GET_LEDGER();
  const isEmailAllowed = LEDGER_COL.EMAIL_ALLOWED - 1;    // Make 0-indexed for arr
  const res = [];

  // Get activity stats in metric and US imperial
  const allStats = convertAndFormatStats_(activity);

  // Transform key labels in Strava to placeholder names in email
  const { metric: metricStats, imperial: imperialStats } = filterEmailValues(allStats);

  // Loop through emails, package member data, then send email
  for (const email of recipients) {
    const entry = getLedgerEntry_(email, ledgerData);

    if (!entry[isEmailAllowed]) continue;   // Only sent to members who consented

    const memberTotalStats = sheetToEmailLabels(entry);  // Get values for post-run email
    const preferredStats = memberTotalStats['USE_METRIC'] ? metricStats : imperialStats;

    // Email report and log response
    res.push(sendPostRunEmail_(email, { ...memberTotalStats, ...preferredStats }));
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

/**
 * Creates post run email from member stats and template,
 * sends it to given email address
 * 
 * @param {string} email  Email address of member
 * @param {string: *} memberStats  Object containing member's information
 * 
 * @returns {string}  Confirmation message
 */
function sendPostRunEmail_(email, memberStats) {
  // Create template to populate
  const template = HtmlService.createTemplateFromFile(POST_RUN_TEMPLATE);
  
  // Add subject line
  template.SUBJECT_LINE = POINTS_EMAIL_SUBJECT_LINE;

  // Get member's system preference to format email
  const useMetric = memberStats['USE_METRIC'];
  template.USE_METRIC = useMetric;

  // Populate member's general stats
  template.FIRST_NAME = memberStats['FIRST_NAME'];
  template.TPOINTS = memberStats['TPOINTS'];
  template.TWEEKS = memberStats['TWEEKS'];
  template.TRUNS = memberStats['TRUNS'];

  // Populate head run details
  template.RUN_DATE = memberStats['RUN_DATE'];
  template.LEVEL = memberStats['LEVEL'];
  template.HEADRUNNERS = memberStats['HEADRUNNERS'];

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
  
  MailApp.sendEmail(
    message = {
      to: email,
      bcc: 'andrey.gonzalez@mail.mcgill.ca',
      name: EMAIL_SENDER_NAME,
      subject: POINTS_EMAIL_SUBJECT_LINE,
      replyTo: MCRUN_EMAIL,
      htmlBody: filledTemplate.getContent(),
    }
  );

  // Log confirmation for the sent email with member stats
  const confirmation = `Stats email sent to ${email} with ${useMetric ? 'metric' : 'imperial'} units.`;
  logAsPL_(confirmation, sendPostRunEmail_.name);
  return confirmation;
}


/**
 * Automatically triggered to send win back email to members whose
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

  // columns (0 indexed)
  const EMAIL_COL = LEDGER_COL.EMAIL - 1;    // const EMAIL_COL = 0;
  const FNAME_COL = LEDGER_COL.FIRST_NAME - 1;   // const FNAME_COL = 2;
  const LAST_RUN_COL = LEDGER_COL.LAST_RUN_DATE - 1;   // const LAST_RUN_COL = 10;

  // make date object for 2 weeks ago
  let dateThreshold = new Date();
  dateThreshold.setDate(dateThreshold.getDate() - 14);

  // get all data entries as 2d array (row, col)
  let allMembers = LEDGER_SHEET.getDataRange().getValues();

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
        sendWinBackEmail_(member[EMAIL_COL], member[FNAME_COL]);
      }
    }
  }
}


/**
 * Creates win back email from member name and template,
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
function sendWinBackEmail_(email, name) {
  // set up email using template
  const template = HtmlService.createTemplateFromFile(WINBACKEMAIL_TEMPLATE);
  template.FIRST_NAME = name;
  let filledTemplate = template.evaluate();

  // send email
  try {
    MailApp.sendEmail(
      message = {
        to: email,
        name: EMAIL_SENDER_NAME,
        subject: WINBACKEMAIL_SUBJECT,
        htmlBody: filledTemplate.getContent()
      }
    );

  } catch (e) {
    Logger.log(e);
  }

  // Log confirmation for the sent email
  Logger.log(`Win-back email sent to ${email}.`);
}
