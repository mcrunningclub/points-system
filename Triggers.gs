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
 * Handler for GET events sent to web app deployment of this script
 * 
 * Checks for authorization, then sets trigger to check for Strava activity
 * for the row specified in the request.
 * 
 * @param {Object} e  GET request object, should contain keys 'key' and 'rowNum
 * @return {TextOutput}  Message indicating status of request
 */
function doGet(e) {
  // 1. Check if access is authorized with key
  if (e.parameter.key !== getSecretWebKey_()) {
    return ContentService.createTextOutput("Unauthorized! Please verify key.");
  }

  // 2. Get 'rowNum' from URL and validate input
  let rowNum = e.parameter.rowNum;
  if (!rowNum || isNaN(rowNum)) {
    return ContentService.createTextOutput("Invalid or missing 'rowNum' parameter.");
  }

  // 3. Parse for row number
  rowNum = parseInt(rowNum, 10);
  Logger.log(`[PL#doGet] Received in 'doGet' row number: ${rowNum}`);

  // 4. Run handler function and return output message
  createStravaTrigger_(rowNum);
  return ContentService.createTextOutput(`Trigger set for row ${rowNum}`);

  /** Helper: get secret key in script properties */
  function getSecretWebKey_() {
    const property = 'WEB_APP_KEY';
    return PropertiesService.getScriptProperties().getProperty(property);
  }
}


/**
 * Creates a trigger to check for Strava activity for a given row, and stores its information
 * in script properties. The property key includes the row number and its values contains the 
 * number of tries, trigger ID, and row number.
 * 
 * @param {integer} row  Row to check for activity for. Defaults to last row
 */
function createStravaTrigger_(row = getValidLastRow_(LOG_SHEET)) {
  const scriptProperties = PropertiesService.getScriptProperties();

  const trigger = ScriptApp.newTrigger(checkForStravaActivities.name)
    .timeBased()
    .everyMinutes(TRIGGER_FREQUENCY)
    .create();

  // Store trigger details using rowNumber as key
  const triggerData = {
    tries: 1,
    triggerId: trigger.getUniqueId(),
    rowNumber: row
  };

  // Label trigger with row number, and log trigger data
  const key = TRIGGER_BASE_ID + row;
  const dataStr = JSON.stringify(triggerData);

  scriptProperties.setProperty(key, dataStr);
  logAsPL_(`Created new trigger '${key}', running every ${TRIGGER_FREQUENCY} min\n${dataStr}`);
}


/**
 * Function called by Strava triggers. Checks for Strava activity for all the rows that currently
 * have active triggers according to script properties, and increments the number of runs for 
 * each one. If max number of tries reached, deletes the trigger.
 */
function checkForStravaActivities() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProps = scriptProperties.getProperties();

  for (let key in allProps) {
    if (!key.startsWith(TRIGGER_BASE_ID)) continue;

    const triggerData = JSON.parse(allProps[key]);
    const { rowNumber, tries, triggerId } = triggerData;
    console.log(`Trigger data`, triggerData);

    if (isStravaFound(rowNumber) && isEmailsSent(rowNumber)) {
      // Activity found during last try, clean up trigger and data in script properties
      cleanUpTrigger(key);
      Logger.log(`✅ Activity found for row ${rowNumber} after ${tries} tries`);
    }
    else if (tries <= MAX_STRAVA_CHECKS) {
      // Activity not found and max tries not reached, check again
      Logger.log(`Strava activity check #${triggerData.tries} for row ${rowNumber}`);
      triggerData.tries++;
      scriptProperties.setProperty(key, JSON.stringify(triggerData));
      Logger.log(`Incremented tries for Strava trigger. Now sending stats email`);
      checkAndSendPostRunEmail(GET_LOG_SHEET(), rowNumber);   // This checks for Strava activity and sends post-run email if success
    }
    else {
      // Max tries reached, clean up trigger and send email notifying activity not found
      cleanUpTrigger(key);
      alertStravaActivityNotFound_(rowNumber, tries);
      Logger.log(`❌ Max tries reached for row ${rowNumber}, sending email and stopping checks`);
    }
  }

  /** Helper: check if Strava activity already logged */
  function isStravaFound(row) {
    const sheet = GET_LOG_SHEET();
    const value = sheet.getRange(row, LOG_COL.STRAVA_ACTIVITY_ID).getValue();
    return value.toString().trim() != '';
  }

  /** Helper: check if post run email already sent */
  function isEmailsSent(row) {
    const sheet = GET_LOG_SHEET();
    const value = sheet.getRange(row, LOG_COL.EMAIL_STATUS).getValue();
    return value.toString().trim() != '';
  }
}

/** 
 * Remove trigger and its property in script properties 
 * 
 * @param {string} propertyKey  Key of the script property corresponding to the trigger to delete
 */
function cleanUpTrigger(propertyKey) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProps = scriptProperties.getProperties();

  const triggerData = allProps[propertyKey];
  const triggerId = triggerData.triggerId;

  if (!deleteTriggerById_(triggerId)) {
    alertTriggerNotFound_(triggerData);
  }
  // Delete property whether trigger is found or not
  scriptProperties.deleteProperty(propertyKey);
}

/** 
 * Delete a trigger given its ID 
 * 
 * @param {string} id  ID of the trigger to delete
 * 
 * @returns {boolean}  True if successfully deleted, otherwise false
 */
function deleteTriggerById_(id) {
  const triggers = ScriptApp.getProjectTriggers();

  for (let trigger of triggers) {
    if (trigger.getUniqueId() === id) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Trigger with id ${id} deleted!`);
      return true;
    }
  }

  // Notify club of unidentified trigger
  console.error(`Unable to find trigger with id #${id}`);
  return false;
}

/**
 * Removes all Strava triggers in ScriptApp.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 16, 2025
 * @update  Nov 16, 2025
 */

function deleteAllStravaTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName == TRIGGER_FUNC) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted Strava trigger with id '${trigger.getUniqueId()}'`);
    }
  });

  Logger.log(`Deleted all Strava triggers successfully!`);
}


/**
 * Sends email to club account saying that a (Strava) trigger was not found and
 * so could not be deleted.
 * 
 * @param {Object} triggerData  Script property value corresponding to the trigger
 */
function alertTriggerNotFound_(triggerData) {
  const triggerId = triggerData.triggerId;

  MailApp.sendEmail({
    to: MCRUN_EMAIL,
    subject: `Trigger not found - Points Ledger Code`,
    body: `
    The script attempted to delete trigger with id ${triggerId} in 'Points Ledger' but was unsuccessful.

    Properties stored following value... Warning: values unrelated to trigger ${triggerId}.
      
    ${JSON.stringify(triggerData)}
      
    Please verify manually, and update properties script if required.
    
    (Updated: June 15, 2025)`.replace(/[ \t]{2,}/g, '').trim(),
  });
}

/**
 * Sends email to club account saying that a Strava activity could not be found.
 * 
 * @param {integer} rowNumber  Row that the activity was supposed to be added to
 * @param {integer} tries  Number of attempts that the script made to find the activity
 */
function alertStravaActivityNotFound_(rowNumber, tries) {
  MailApp.sendEmail({
    to: MCRUN_EMAIL,
    subject: `Strava Activity Not Found - Points Ledger Code`,
    body: `
    The script attempted ${tries} times to find a Strava activity for row ${rowNumber} in 'Points Ledger' unsuccessfully.
    
    Please verify manually, and send post-run email to attendees once found.

    (Updated: June 15, 2025)`.replace(/[ \t]{2,}/g, '').trim(),
  });
}

