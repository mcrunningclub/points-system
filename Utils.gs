
/**
 * Gets the timezone of the script
 * 
 * @return {string} Time zone
 */
function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}

/**
 * Gets the email of the user accessing the script
 * 
 * @return {string}  Email address
 */
function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}

/**
 * Tries to find a file in Google Drive by its name
 * 
 * May have an error if there are no results?
 * 
 * @param {string} name  Name of the file to find
 * @return {File}  The first result of the search
 */
function getFileByName_(name) {
  return DriveApp.searchFiles(`title contains '${name}'`).next();
}

/**
 * Tries to find a file in Google Drive by its ID
 * 
 * May have an error if there are no results?
 * 
 * @param {string} id  ID of the file to find
 * @return {File}  The first result of the search
 */
function getFileById_(id) {
  return DriveApp.getFileById(id);
}

/**
 * Creates log in the console with specific formatting
 * 
 * @param {string} msg  The message to log
 * @param {string} [funcName]  Name of the function returning the message, if applicable. Defaults to ""
 * @param {boolean} ][useLogger]  Whether to use the logger (true) or console.log (false). Defaults to true.
 */
function logAsPL_(msg, funcName = "", useLogger = true) {
  const message = `[PL#${funcName}] ${msg}`;
  useLogger ? Logger.log(message) : console.log(message);
}


/**
 * Find row index of last entry, starting from bottom using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Target sheet.
 * @return {integer}  Returns 1-index of last row in `sheet`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sept 1, 2024
 * @update  May 25, 2025
 */

function getValidLastRow_(sheet) {
  const startRow = 1;   // Do not skip header row here
  const numRow = sheet.getLastRow();

  // Fetch all values
  const values = sheet.getSheetValues(startRow, 1, numRow, 1);
  let lastRow = values.length;

  // Loop through the values in reverse order
  while (values[lastRow - 1][0] === "") {
    lastRow--;
  }

  return lastRow;
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


/**
 * Fill template string with data object
 * @author Martin Hawksey
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
 * Simple logging of multi-line message. Improves readability in code.
 * @param {string} msg  The message(s) to log
 */
const prettyLog_ = (...msg) => console.log(msg.join('\n'));


/**
 * Convert a Date timestamp to a Unix Epoch timestamp.
 * 
 * @param {Date} timestamp  Timestamp to convert.
 * @return {integer}  Number of seconds elapsed since January 1, 1970.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @date  Dec 1, 2024
 * @update  Dec 1, 2024
 */

function getUnixEpochTimestamp_(timestamp) {
  return Math.floor(timestamp.getTime() / 1000);
}