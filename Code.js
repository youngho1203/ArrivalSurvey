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
// 도착 학생 목록 ( Response List )
const listsSheet = ws.getSheetByName("Response List");
// 방 설정 정보
const configSheet = ws.getSheetByName("Config");
// 입사 학생 정보
const dataSheet = ws.getSheetByName("Data");
// Invoice Template
const templateSheet = ws.getSheetByName("Template");
// 현황 List Unique ID : Arrival Survey 를 입사생이 진행하면, 그 결과를 기숙사 현황 List 에도 같이 등록하기 위한 현황 List ID
// 광개토관 기숙사 : 0. 광개토관 기숙사 현황 ID
// 외부 기숙사 : 
const RESIDENCE_LIST_ID = '19eaIfuIvedUCf0l2V_ZbcP6hfFqmFNWI2sfW1OVWUig';
// 현황 SpreadSheet ( 사전에 만들어져 있어야 한다. )
const residenceListSheet = SpreadsheetApp.openById(RESIDENCE_LIST_ID);
// 현황 SpreadSheet 현재 진행 Tab Name ( 사전에 설정되어 있어야 한다. )
const currentListName = configSheet.getRange("T2").getValue();

// 허용 입사 학생 총 수
const numberOfData = dataSheet.getLastRow() - 1;
// 입사 가능 방 총 수
const availableRooms = configSheet.getLastRow() -1;
// configSheet 에서 nextRoomCode Column number
const nextRoomCodeColumn = 8;
// 입실 가능한 방이 꽉 찼을 때 
const FULL_ROOMS = "FULL";

/**
 * @TODO : 현황 LIST 에 자동으로 행 입력.
 * @TODO : 생일 Data 추가
 * @TODO : 하나의 BED 에 중복 배정 방지 Check 도입
 * @TODO : 배정 진행시 이 빠진 Bed 를 확인하여 Next 보다 우선 설정.
 * @TODO : Data 명단에 없는 학생 등록을 진행할 때 처리 ( NOT FOUND 발생시 InsertArrivalSurvey 로 다시 실시 ?????, 실제 학번을 가지고 있는 학생인지 어떻게 확인???? )
 */
/**
 * Arrival Survey 가 등록되면 실행된다.
 * @param {Object} survey event object
 */
function setInitialValue(e) {
  try {
    var range = e.range.offset(0,1, 1, 1);
    var studentId = range.getValue();
    if(deDupeCheck(studentId) !== undefined){
      // console.log(deDupeCheck(studentId));
      throw new Error("[" + studentId + "] is Aleady CheckIn");
    }
    var studentInfo = getStudentInfo(studentId);
    if(studentInfo == undefined) {
      throw new Error("Can Not Find Your StudentId [" + studentId + "]");
    } 
    //
    doBuild(range, studentInfo, 'A');
    //
    // 현황 List 에 내용을 추가한다.
    //
    var values = e.range.getValues()[0];
    studentInfo.email = values[2];
    studentInfo.phone = values[3];
    // console.log("Now add to 현황 List : ", studentInfo);
    //
    // 현황 List 는 미리 준비되어 있어야 한다.
    appendResidence(studentInfo);
  }
  catch(e) {
    range.clearContent();
    range.offset(0, 6, 1, 1).setValue(e);
    range.offset(0,-1,1,8).setBackground("Orange");
  }
}

/**
 * dedupe check for duplication checkin
 * Survey Response List 에서만 확인한다.
 */
function deDupeCheck(studentId) {
  // lastRow 는 지금 진행하고 있는 것. 바로 직전까지만 처리
  var lastRow = listsSheet.getLastRow() -1;
  var range = listsSheet.getRange("B2:B" + lastRow);
  return range.getValues().find( id => { return id[0] === studentId });
}

/**
 * nextBed 가 앞으로 진행되고 있는데 어떤 이유에서든( 수동 배정 이동 ) 그 앞에 빠진 침대가 있으면 먼저 그 침대에 배정한다.
 * 현황 List 에서 확인한다.
 * @param residenceType
 */
function findSkipBed(residenceType) {
  // [nextCode, firstCode] array
  var code_range = configSheet.getRange(residenceType, nextRoomCodeColumn, 1, 2).getValues()[0];
  currentList = residenceListSheet.getSheetByName(currentListName);
  // residenceType 별 row range 를 구한다.
  var startRow;
  var lastRow;
  var totalLastRow = currentList.getLastRow();
  currentList.getRange("B3:C" + totalLastRow).getValues().forEach((value, index) => {
    if(value.join('') == code_range[1]) {
      startRow = index + 3;
    }
    else if(value.join('') == code_range[0]) {
      // nextCode 바로 전 까지만 확인 
      lastRow = index + 2;
    }
  });

  // 학번이 공란이 것을 확인한다.
  var skipBedCode = '';
  currentList.getRange("A" + startRow + ":E" + lastRow).getValues().forEach((value,index) => { 
    if(value[4] == '' && skipBedCode == '') {
      // 해당 bed 정보를 return
      // 순번은 1부터 순차적으로 증가하여야 한다.
      skipBedCode = currentList.getRange(value[0] + 2, 2, 1, 2).getValues()[0].join('');
    }
  });
  return skipBedCode;
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
    //
    // Residence data 를 update 한다.
    //
    updateResidence(studentInfo);
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
  //
  if(studentInfo.isPreAssigned) {
    // 수동으로 설정해 놓았으면 처리하지 않는다.
  }
  else {
    if(nextRoomCode === FULL_ROOMS) {
      throw new Error("방이 모두 찾습니다. 더 이상 배정을 할 수 없습니다.");
    }
    else {
      // skipBed 가 존재하면, skipBed 로 설정한다.
      var skipBed = findSkipBed(row);
      if(!isCellEmpty(skipBed)) {
        studentInfo.assignedRoom = skipBed;
        studentInfo.isPreAssigned = true;
      }
      else {      
        //
        // @todo make sure synchronized block
        //
        nextRoomCode = configSheet.getRange(row, nextRoomCodeColumn).getValue();      
        studentInfo.assignedRoom = nextRoomCode;
        updateNextRoomNumberCode(row, studentInfo);
      }
    }
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
  var lastRoomCode = configSheet.getRange(row, (nextRoomCodeColumn + 2) ).getValue();
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
  let residenceInfo = configSheet.getRange(residenceType, 7, 1, 9).getValues()[0];
  /**
   * 'Residence Type',
   * 'Next Assigned Room Code',	
   * 'First Room Code',
   * 'Last Room Code',
   * 'numberOfMonth',
   * 'Residence Period',	
   * 'Payment Peroid',
   * 'defaultFee',
   * 'alias'
   */
  // Residence Period
  let residencePeriod = residenceInfo[5].split('~');
  return {
    'type': residenceInfo[0],
    'numberOfMonth': residenceInfo[4],
    'availableDate': residencePeriod[0],
    'dueDate' : residencePeriod[1],
    'paymentPeriod':residenceInfo[6], // 
    'defaultFee': residenceInfo[7], // 무료 학생의 기본 기숙사 비
    'aliasPattern': residenceInfo[8] // 기숙사 주소 alias Pattern
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
      'aliasPattern': '', // 기숙사 이름 alias pattern
      'phone':'',
      'email':''
      };
  }
  return undefined;
}

/**
 * ResidenceList 에 방 배정된 학생 정보를 등록한다.
 * @param {Object} studentInfo
 */
function appendResidence(studentInfo) {
  rowData = [[
    false, //D : 퇴사 ( CheckBox ) : 퇴사시 Check 하면 해당 Row 를 퇴사한 것으로 변경한다.
    studentInfo.studentId, // E : 학번
    studentInfo.name, // F : 이름
    studentInfo.nationality, // G : 국적
    studentInfo.gender, // H : 성별
    '', // I : 생년월일 : Cell 자료 서식이 반드시 '날짜' 형식 이어야 한다. ( 거주 증명서 발행시 생일을 'YYYY-MM-DD' 형식으로 출력하기 위함 )
    '', // J : 입사 보고일
    '', // K : 납부
    '', // L : 메디컬
    '', // M : Gen Mode 
    '', // N : 시설 점검표
    '', // O : 연장여부
    studentInfo.phone, // P : 핸드폰
    studentInfo.email // Q : 이메일
  ]];
  // console.log(studentInfo);
  var lastLow = residenceListSheet.getLastRow();
  residenceListSheet.getRange("B3:C" + lastLow).getValues().forEach((array, index) => {
    if(array.join('') == studentInfo.assignedRoom) {
      residenceListSheet.getRange("D" + (index + 3) + ":Q" + (index + 3)).setValues(rowData);
    }
  })
}

/**
 * ResidenceList 에 student 를 변경한다.
 * @param {Object} studentInfo
 */
function updateResidence(studentInfo) {
  // @todo need implements
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

// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}