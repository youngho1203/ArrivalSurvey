
/**
 * assignedRoom 정보에서 dormitory 주소, 거주기한, fee 를 얻음.
 */
function _setDormitoryInfo(residenceType, studentInfo) {
  // 기숙사 거주 유형별 정보
  const residenceInfo = getResidenceInfo(residenceType);
  //
  // 침대는 최대 9개 미만 ( 알파벳 한자리 )
  var str_length = studentInfo.assignedRoom.length;
  var roomNumber = studentInfo.assignedRoom.substring(0,str_length - 1);
  configSheet.getRange("B2:B" + (2+ availableRooms)).getValues().forEach((room, index) => {
    if(room == roomNumber){
      var roomInfo = configSheet.getRange("A" + (2 + index) + ":C" + (2 + index)).getValues();
      roomInfo = roomInfo[0]
      /** 
       * roomInfo array
       * 'Domitory Name',	
       * 'Available Rooms',	
       * 'Beds'
      */
      if(studentInfo.isFree) {
        studentInfo.dormFee = residenceInfo.defaultFee;
      }
      else {
        studentInfo.dormFee = residenceInfo.dormFee;
      }
      // 
      studentInfo.dormName = roomInfo[0];
      studentInfo.availableDate = residenceInfo.availableDate;
      studentInfo.dueDate = residenceInfo.dueDate;
      studentInfo.paymentPeriod = residenceInfo.paymentPeriod;
    }
  })
}


function _getResidenceInfo(residenceType) {
  let residenceInfo = configSheet.getRange(residenceType, 4, 1, 6).getValues();
  residenceInfo = residenceInfo[0];
console.log('residenceInfo', residenceInfo);
}
