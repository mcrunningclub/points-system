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