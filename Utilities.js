/**
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
 * Creates the menu item "Manual Work" for manual build of student invoice
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('방 배정 수정')
      .addItem('수정 Invoice 발행', 'showDialog')
      .addToUi();
}

function getDataFromFormSubmit(form) {
  buildInvoidByManual(form.studentId, form.code);  
}

function showDialog() {
  // Display a modal dialog box with custom HtmlService content.
    var dialog = HtmlService.createHtmlOutputFromFile("Dialog.html").setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(dialog, '변경할 내역을 입력하세요');
}

/**
 * 'yyyy-mm-dd' date String
 */
function _getNowDateISOFormattedString(){
  return _getISOTimeZoneCorrectedDateString(new Date());
}

/**
 * javascript toISOString timezone treatment
 */
function _getISOTimeZoneCorrectedDateString(date) {
  // timezone offset 처리 
  var tzoffset = date.getTimezoneOffset() * 60000; //offset in milliseconds
  return (new Date(date.getTime() - tzoffset)).toISOString().substring(0, 10);
}
