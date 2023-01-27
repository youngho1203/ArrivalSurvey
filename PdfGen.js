/**
Copyright 2022 Google LLC

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
//
// Application constants
const APP_TITLE = 'Generate Invoice PDF';
const OUTPUT_FOLDER_NAME = "Invoce of Arrival Student PDFs";

/**
* Resets the template sheet by clearing out studentInfo data.
* You use this to prepare for the next iteration or to view blank
* the template for design.
*/
function clearTemplateSheet(sheet) {
  // Clears existing data from the template.
  const rngClear = sheet.getRangeList(['A5', 'B6:B11', 'B13', 'B14']).getRanges()
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
}

/**
 * Creates a PDF for the customer given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created : studentId_roomNumberCode
 * @return {file object} PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName) {
  // const fr = 0, fc = 0, lc = 9, lr = 27;
  // const fr = 0, fc = 0, lc = 0, lr = 29;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=a4&" +          // paper A4 
    "fzr=true&" +         // do not repeat row headers
    "portrait=false&" +   // landscape
    "fitw=true&" +        // fit to page width
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.30&" +
    "bottom_margin=0.00&" +
    "left_margin=0.60&" +
    "right_margin=0.00&" +
    "sheetnames=false&" +
    "pagenum=false&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId();
    /** 
     * + "&r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;
     */
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');
  // pdf file saved folder
  const pdfFolder = getFolderByName_(OUTPUT_FOLDER_NAME);
  // Gets the folder in Drive where the PDFs are stored.
  return pdfFolder.createFile(blob);
}

/**
 * Returns a Google Drive folder in the same location 
 * in Drive where the spreadsheet is located. First, it checks if the folder
 * already exists and returns that folder. If the folder doesn't already
 * exist, the script creates a new one. The folder's name is set by the
 * "OUTPUT_FOLDER_NAME" variable from the Code.gs file.
 *
 * @param {string} folderName - Name of the Drive folder. 
 * @return {object} Google Drive Folder
 */
function getFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  //
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();
  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by ${APP_TITLE} application to store PDF output files`);
}