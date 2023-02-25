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
const ws = SpreadsheetApp.getActiveSpreadsheet();
// 도착 학생 목록
const listsSheet = ws.getSheetByName("Response List");
// 방 설정 정보
const configSheet = ws.getSheetByName("Config");
// 입사 학생 정보
const dataSheet = ws.getSheetByName("Data");
// Invoice Template
const templateSheet = ws.getSheetByName("Template");
// 허용 입사 학생 총 수
const numberOfData = dataSheet.getLastRow() - 1;
// 입사 가능 방 총 수
const availableRooms = configSheet.getLastRow() -1;
// configSheet 에서 nextRoomCode Column number
const nextRoomCodeColumn = 8;
// 입실 가능한 방이 꽉 찼을 때 
const FULL_ROOMS = "FULL";

/**
 * @TODO : Response List : Name column add 
 */
/**
 * Arrival Survey 가 등록되면 실행된다.
 * @param {Object} survey event object
 */
function setInitialValue(e) {
  try {
    var range = e.range.offset(0,1, 1, 1);
    var studentId = range.getValue();
    if(deDupeCheck(studentId)) {
      throw new Error("[" + studentId + "] is Aleady CheckIn");
    }
    var studentInfo = getStudentInfo(studentId);
    if(studentInfo == undefined) {
      throw new Error("Can Not Find Your StudentId [" + studentId + "]");
    } 
    //
    doBuild(range, studentInfo, 'A');
  }
  catch(e) {
    range.offset(0, 6, 1, 1).setValue(e);
  }
}

/**
 * dedupe check for duplication checkin
 */
function deDupeCheck(studentId) {
  // lastRow 는 지금 진행하고 있는 것. 바로 직전까지만 처리
  var lastRow = listsSheet.getLastRow() -1;
  var range = listsSheet.getRange("B2:B" + lastRow);
  return range.getValues().find( id => { return id[0] === studentId });
}

/**
 * 기 발행된 invoice 에서 방을 변경하기 위하여 Menu 에서 수동으로 변경 진행할 때 
 * @param {Number} 수정하고자 하는 학생 학번
 * @param {String} 수정하고자 하는 Code
 */
function buildInvoidByManual(studentId, roomCode){
  //
  var lastRow = listsSheet.getLastRow() + 1;
  var range = listsSheet.getRange(lastRow, 1);
  range.setValue(new Date());  
  range = range.offset(0, 1, 1, 1);
  range.setValue(studentId);
  
  try {
    var studentInfo = getStudentInfo(studentId);
    if(studentInfo == undefined) {
      throw new Error("Can Not Find Your StudentId [" + studentId + "]");
    }    
    studentInfo.assignedRoom = roomCode;
    studentInfo.isPreAssigned = true;
    //
    var modified_pdf_url = doBuild(range, studentInfo, 'M');
    //
    // DataSheet 에 학생의 AssignedRoom 에 Manual 설정값을 기록한다. 
    // ( findNextCode 로직을 동일하게 유지시킨다. ) 
    //
    dataSheet.getRange("A2:A" + (1 + numberOfData)).getValues().forEach((value, index) => {
      if(value[0] == studentId){
        dataSheet.getRange(index + 2, 7).setValue(roomCode);
      }
    });
    //
    // 앞서서 Survey 진행한 정보에서 학생의 new roomCode, modified_pdf_url 을 설정한다.
    //
    listsSheet.getRange("B2:B" + (lastRow -1)).getValues().forEach((value, index) => {
      if(value == studentId){
        // timestamp
        listsSheet.getRange(index + 2, 1, 1, 1).setValue(new Date());
        // roomCode
        listsSheet.getRange(index + 2, 5, 1, 1).setValue(roomCode);
        // doomfee
        listsSheet.getRange(index + 2, 6, 1, 1).setValue(studentInfo.dormFee);
        // modifiedPdfUrl
        listsSheet.getRange(index + 2, 8, 1, 1).setValue(modified_pdf_url);
        listsSheet.getRange(index + 2, 1, 1, 8).setBackground("#e0e0e0");
      }
    });
    // M 을 지운다.
    listsSheet.deleteRow(lastRow);
  }
  catch(e) {
    range.offset(0, 6, 1, 1).setValue(e);
  }
}

/**
 * main build
 * @return invoice_url
 */
function doBuild(range, studentInfo, genType) {
  //
  setRoomNumberCode(studentInfo);
  // 
  range.offset(0, 3, 1, 1).setValue(studentInfo.assignedRoom);
  range.offset(0, 4, 1, 1).setValue(studentInfo.dormFee);
  // generation mode 에 따른 marker
  range.offset(0, 5, 1, 1).setValue(genType);
  // build invoice now
  var invoice_url = buildInvoicePdf(studentInfo);
  range.offset(0, 6, 1, 1).setValue(invoice_url);
  return invoice_url;
}

/**
 * Invoice Template 를 읽어서 pdf file 을 생성한다. 
 * @param {Object} studentInfo
 * @return pdf url or error
 */
function buildInvoicePdf(studentInfo) {
  try {
    var url = createInvoiceForStudent(studentInfo, templateSheet, ws.getId());
    // toast is working on Manual Mode only.
    ws.toast("방 변경 Invoice 를 생성하였습니다.", '', 2);
    return url;
  }
  catch(e) {
    return e;
  }
}

/**
 * @param {Object} studentInfo
 */
function setRoomNumberCode(studentInfo) {
  var gender= studentInfo.gender;
  var isExchangeStudent = studentInfo.isExchangeStudent;
  // next roomCode 는 ConfigSheet 에 기록하여 놓았던 것을 읽는다. ( ID Column 이다. )
  // row 는 residence type 이다.
  var nextRoomCode, row; 
  if(gender.startsWith('F')) {
    // female
    if(isExchangeStudent) {
      row = 2;
    }
    else {
      row = 4;
    }
  }
  else {
    // male
    if(isExchangeStudent) {
      row = 3;
    }
    else {
      row = 5;
    }
  }
  if(studentInfo.isPreAssigned) {
    // 수동으로 설정해 놓았으면 처리하지 않는다.
  }
  else {
    //
    // @todo make sure synchronized block
    //
    nextRoomCode = configSheet.getRange(row, nextRoomCodeColumn).getValue();
    if(nextRoomCode === FULL_ROOMS) {
      throw new Error("방이 모두 찾습니다. 더 이상 배정을 할 수 없습니다.");
    }
    else {
      studentInfo.assignedRoom = nextRoomCode;
    }
    updateNextRoomNumberCode(row, studentInfo);
  }
  setDormitoryInfo(row, studentInfo);
}

/**
 * assignedRoom 정보에서 dormitory 주소, 거주기한, fee 를 얻음.
 */
function setDormitoryInfo(residenceType, studentInfo) {
  // 기숙사 거주 유형별 정보
  const residenceInfo = getResidenceInfo(residenceType);
  //
  // 침대는 최대 9개 미만 ( 알파벳 한자리 )
  var str_length = studentInfo.assignedRoom.length;
  var roomNumber = studentInfo.assignedRoom.substring(0,str_length - 1);
  configSheet.getRange("B2:B" + (1+ availableRooms)).getValues().forEach((room, index) => {
    // white space 제거, human error 를 방지
    if(room[0].toString().replace(/\s/g, "") == roomNumber.replace(/\s/g, "")){
      var roomInfo = configSheet.getRange("A" + (2 + index) + ":D" + (2 + index)).getValues()[0];
      /** 
       * roomInfo array
       * 'Domitory Name',	
       * 'Available Rooms',	
       * 'Beds'
       * 'DomFee per Month'
      */
      if(studentInfo.isFree) {
        studentInfo.dormFee = residenceInfo.defaultFee;
      }
      else {
        //
        // 기본 요금 + 각 dormitory 의 단위 요금 * 거주 기간 = 기숙사 비
        //        
        studentInfo.dormFee = residenceInfo.defaultFee + residenceInfo.numberOfMonth * roomInfo[3];
      }
      // 
      studentInfo.dormName = roomInfo[0];
      studentInfo.availableDate = residenceInfo.availableDate;
      studentInfo.dueDate = residenceInfo.dueDate;
      studentInfo.paymentPeriod = residenceInfo.paymentPeriod;
      studentInfo.aliasPattern = residenceInfo.aliasPattern;
    }
  });
}

/**
 * ConfigSheet 에 Next RoomNumberCode 를 update 한다.
 * @param row number
 * @param {Object} studentInfo
 */
function updateNextRoomNumberCode(row, studentInfo) {
  //
  var roomCode = studentInfo.assignedRoom;
  var roomNumber = roomCode.substring(0, roomCode.length -1);
  var bedCode =   roomCode.substring(roomCode.length -1);

  // next 침대
  var nextRoomCode;
  if(isLastRoom(row, studentInfo.assignedRoom)){
    nextRoomCode = FULL_ROOMS;
  }
  else {
   nextRoomCode = findNextCode(roomNumber, bedCode);
    // 미리 할당된 침대인지 dataSheet 확인
    dataSheet.getRange("G2:G" + (1+ numberOfData)).getValues().forEach(value => {
      if(value == nextRoomCode){
        // 이미 할당된 침대이면 다음 침대
        roomNumber = nextRoomCode.substring(0, nextRoomCode.length -1);
        bedCode =   nextRoomCode.substring(nextRoomCode.length -1, nextRoomCode.length);    
        nextRoomCode = findNextCode(roomNumber, bedCode);
      }
    });
  }
  configSheet.getRange(row, nextRoomCodeColumn).setValue(nextRoomCode);
}

/**
 * 더 배정 가능한 방이 있는지 여부 확인
 * @param {Number} row 
 * @param {Object} nextRoomCode 
 */
function isLastRoom(row, nextRoomCode) {
  var lastRoomCode = configSheet.getRange(row, (nextRoomCodeColumn + 1) ).getValue();
  return (lastRoomCode === nextRoomCode);
}

/**
 * @return {String} 
 */
function findNextCode(roomNumber, bedCode) { 
  var nextRoom, nextCode; 
  configSheet.getRange("B2:B" + (1+ availableRooms)).getValues().forEach((room, index) => {
    if(room == roomNumber){
      // index 는 0 부터, 따라서 C2 부터 
      var bedArray = configSheet.getRange("C" + (2 + index) ).getValue().split(',');

      if(bedArray.indexOf(bedCode) == (bedArray.length -1)) {
        // next room, first bed
        nextCode = bedArray[0];
        nextRoom = configSheet.getRange("B" + (2 + index + 1)).getValue();
      }
      else {
        // same room, next bed
        nextCode = bedArray[bedArray.indexOf(bedCode) + 1];
        nextRoom = room;
      }
    }
  });
  return nextRoom + nextCode;
}

/**
 * residenceType 에 따라서 ResidenceInfo 를 구한다.
 * @param {array} residenceType
 * @return {Object} residenceInfo
 */
function getResidenceInfo(residenceType) {
  // 'G' column 부터 8개 column
  let residenceInfo = configSheet.getRange(residenceType, 7, 1, 8).getValues()[0];
  /**
   * 'Residence Type',
   * 'Next Assigned Room Code',	
   * 'Last Room Code',
   * 'numberOfMonth',
   * 'Residence Period',	
   * 'Payment Peroid',
   * 'defaultFee',
   * 'alias'
   */
  // Residence Period
  let residencePeriod = residenceInfo[4].split('~');
  return {
    'type': residenceInfo[0],
    'numberOfMonth': residenceInfo[3],
    'availableDate': residencePeriod[0],
    'dueDate' : residencePeriod[1],
    'paymentPeriod':residenceInfo[5], // 
    'defaultFee': residenceInfo[6], // 무료 학생의 기본 기숙사 비
    'aliasPattern': residenceInfo[7] // 기숙사 주소 alias Pattern
  };
}

/**
 * DataSheet 에서 matching 되는 학생 정보를 찾는다. 
 */
function getStudentInfo(studentId) {
  var studentData;
  dataSheet.getRange(2,1,numberOfData).getValues().forEach((id, index) => {
    if(id == studentId) {
      studentData = dataSheet.getRange(index + 2, 1, 1, 7).getValues()[0];
    }
  });
  if(studentData){
    var isAssigned = studentData[6] == '' ? false : true;
    return { 
      'studentId':studentData[0], 
      'name':studentData[1], 
      'nationality':studentData[2], 
      'gender': studentData[3], 
      'isFree':studentData[4], 
      'isExchangeStudent':studentData[5], 
      'assignedRoom':studentData[6], // 배정된 방
      'isPreAssigned': isAssigned,
      'dormName': '', // Dorm Name
      'dormFee':-1, // 기숙사 비
      'deposit': 0, // deposit money
      'availableDate': '', // 거주 가능 시작 날짜
      'dueDate': '', // 거주 종료 날짜
      'address' :'', // 기숙사 방 주소
      'paymentPeriod':'', // 기숙사비 납부 일정
      'aliasPattern': '' // 기숙사 이름 alias pattern
      };
  }
  return undefined;
}

/**
 * INVOICE PDF File Name Pattern = StudentId_RoomNumberCode.pdf
 * @param {Object} studentInfo
 * @param {Object} templateSheet for invoice
 * @param {String} this spreadsheet id  
 * @return {String} created PDF url
 */
function createInvoiceForStudent(studentInfo, sheet, ssId) {

  // Clears existing data from the template.
  clearTemplateSheet(sheet);
  // Sets values in the template.
  sheet.getRange('A5').setValue(studentInfo.dormName);
  sheet.getRange('B6').setValue(studentInfo.studentId);
  sheet.getRange('B7').setValue(studentInfo.name + ' / ' + studentInfo.gender);
  sheet.getRange('B8').setValue(studentInfo.nationality);
  // @see ConfigSheet Alias Pattern
  var aliasPattern = studentInfo.aliasPattern;
  var roomNumber = studentInfo.assignedRoom.substring(0, studentInfo.assignedRoom.length -1);
  var bedCode =   studentInfo.assignedRoom.substring(studentInfo.assignedRoom.length -1);
  var alias = aliasPattern.replace("\{\{ROOM\}\}", roomNumber).replace("\{\{CODE\}\}", bedCode);
  sheet.getRange('B9').setValue(alias); // dorm address alias
  sheet.getRange('B10').setValue(studentInfo.dormFee);
  sheet.getRange('B11').setValue(studentInfo.deposit);
  sheet.getRange('B13').setValue(studentInfo.paymentPeriod);
  sheet.getRange('B14').setValue(studentInfo.availableDate + ' ~ ' + studentInfo.dueDate); // 거주기간

  // Cleans up and creates PDF.
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf

  var pdfFileName = studentInfo.studentId + '_' + studentInfo.assignedRoom;
  // console.log('createInvoiceForStudent', pdfFileName);
  const pdf = createPDF(ssId, sheet, pdfFileName);

  return pdf.getUrl();
}
