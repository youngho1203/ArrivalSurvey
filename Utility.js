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
  console.log('getDataFromFormSubmit', form);
  buildInvoidByManual(form.studentId, form.code);  
}
function showDialog() {
  // Display a modal dialog box with custom HtmlService content.
    var dialog = HtmlService.createHtmlOutputFromFile("Dialog.html");
  SpreadsheetApp.getUi().showModalDialog(dialog, '변경할 학번, 방 Code 를 입력하세요');
}
