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

const CALENDAR_STORE = SCRIPT_PROPERTY.calendarTriggers;
const TRIGGER_FUNC = findAndStoreStravaActivity.name;

/**
 * Add time-based trigger using event information from Calendar.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) + ChatGPT
 * @date  Apr 15, 2025
 * @update  Apr 17, 2025
 */

function createAndStoreTrigger_(event) {
  const props = GET_PROP_STORE_();
  const stored = JSON.parse(props.getProperty(CALENDAR_STORE) || "{}");

  const offset = 60 * 60 * 1000;
  const startTime = new Date(event.getStartTime().getTime() + offset);

  // Only add trigger if new
  if (isExistingTrigger_(startTime, stored)) return;

  const trigger = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .at(startTime)
    .create();

  stored[trigger.getUniqueId()] = startTime.toISOString();

  // Store updated calendar triggers
  props.setProperty(CALENDAR_STORE, JSON.stringify(stored));
  Logger.log(`Trigger created and stored for "${event.getTitle()}" at ${startTime}`);

  // Helper function
  function isExistingTrigger_(time, stored) {
    const triggerTimes = Object.values(stored);
    return (time in triggerTimes);
  }
}

function createNewStravaTrigger(thisHour, thisMinute) {
  const triggerId = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased(thisHour)
    .nearMinute(thisMinute)
    .create();
  
  Logger.log(`Created new trigger (${triggerId}) for ${thisHour}:${thisMinute}`);
  return triggerId;
}

function testRepeatTrigger() {
  // Trigger every 6 hours.
  ScriptApp.newTrigger('myFunction')
    .timeBased()
    .everyHours(6)
    .create();
  // Trigger every Monday at 09:00.
  ScriptApp.newTrigger('myFunction')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).nearMinute()
    .atHour(9)
    .create();
}
