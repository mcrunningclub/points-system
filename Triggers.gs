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

const TRIGGER_FUNC = runStravaChecker.name;
const TRIGGER_BASE_ID = 'stravaTriggerRow';
const STRAVA_CHECK_MAX_TRIES = 3;
const TRIGGER_FREQUENCE = 30;  // Minutes

function createNewStravaTrigger(row = getValidLastRow_(LOG_SHEET)) {
  const scriptProperties = PropertiesService.getScriptProperties();

  const trigger = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .everyMinutes(TRIGGER_FREQUENCE)
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
  Logger.log(`Created new trigger '${key}', running every ${TRIGGER_FREQUENCE} min\n${dataStr}`);
}


// This function will be repeatedly called by the trigger
function runStravaChecker() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProps = scriptProperties.getProperties();

  for (let key in allProps) {
    if (!key.startsWith(TRIGGER_BASE_ID)) continue;

    const triggerData = JSON.parse(allProps[key]);
    const { rowNumber, tries, triggerId } = triggerData;

    if (isStravaFound(rowNumber)) {
      // If found, clean up trigger and data in script properties
      cleanUpTrigger(key, triggerId);
      Logger.log(`✅ Activity found for row ${rowNumber} after ${tries} tries`);
    }
    else if (tries <= STRAVA_CHECK_MAX_TRIES) {
      // Limit not reach, check again and increment 'tries'
      incrementTries(key, triggerData);
      sendStatsEmail();   // This checks for Strava activity and sends post-run email if success
    }
    else {
      // Send email notification if limit is reached
      cleanUpTrigger(key, triggerId);
      alertStravaNotFound_(rowNumber, tries);
      Logger.log(`❌ Max tries reached for row ${rowNumber}, sending email and stopping checks`);
    }
  }

  /** Helper: check if Strava activity already logged */
  function isStravaFound(row) {
    const sheet = GET_LOG_SHEET_();
    const value = sheet.getRange(row, LOG_INDEX.STRAVA_ACTIVITY_ID).getValue();
    return value.toString().trim() != '';
  }

  /** Helper: increment tries and log data */
  function incrementTries(key, triggerData) {
    Logger.log(`Strava activity check #${triggerData.tries} for row ${triggerData.rowNumber}`);
    triggerData.tries++;
    scriptProperties.setProperty(key, JSON.stringify(triggerData));
  }

  /** Helper: remove trigger and data in script properties */
  function cleanUpTrigger(key, triggerId) {
    deleteTriggerById(triggerId);
    scriptProperties.deleteProperty(key);
  }

  /** Helper: delete a trigger by ID */
  function deleteTriggerById(triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    let isFound = false;

    for (let trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        isFound = true;
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
    // Log success or throw error if not found
    const raiseError = () => { throw new Error(`⚠️ Trigger with id ${triggerId} not found`) }
    isFound ? Logger.log(`Trigger with id ${triggerId} deleted!`) : raiseError();
  }
}

function alertStravaNotFound_(rowNumber, tries) {
  MailApp.sendEmail({
    to: MCRUN_EMAIL,
    subject: `Strava Activity Not Found - Row #${rowNumber}`,
    body: `The script attempted ${tries} times to find a Strava activity for row ${rowNumber}, but it was not found.
    
    Please verify manually, and send post-run email once found.`
  });
}

