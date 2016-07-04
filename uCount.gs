/*
##########################
# uCount stats retriever #
##########################

Timezone offset management courtesy of "cryo" (http://stackoverflow.com/users/304185/cryo).
Original post: http://stackoverflow.com/questions/1091372/getting-the-clients-timezone-in-javascript

*/

var sheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');


//MAIN FUNCTION
function launchFollowersRetrievalEngine() {
  
  //array from 0 to 5
  var users = ['University', 'projects', 'Events', 'tools', 'cm', 'Platform'];
  
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
  i = sheetData.getLastRow();
  
  // i value checker
  //Browser.msgBox (i);
  
  //set sydate
  var currentDate = new Date(); 
  
  
  var offset = currentDate.getTimezoneOffset()
  offset = ((offset<0? '+':'-')+ // Note the reversed sign!
            pad(parseInt(Math.abs(offset/60)), 2)+
            pad(Math.abs(offset%60), 2))

  var datetime = addZero(currentDate.getDate()) + "/" + addZero(currentDate.getMonth()+1)  + "/" + currentDate.getFullYear() +
                "  " + addZero(currentDate.getHours()) + ":" + addZero(currentDate.getMinutes()) + ":" + addZero(currentDate.getSeconds()) +
                "  UTC " + offset;

  var cell1 = sheetData.getRange(i+1,1).setValue(datetime);
  
  
  //alternate set date full  
  //var formattedDate = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy T HH:mm:ss z");
  //var cell1 = sheet.getRange(i+1,1).setValue(formattedDate);
  
  //extract followers and count directly in the for loop the total count of followers
  //set itemsNo variable to count the valid accounts
  var itemsCount = 0;
  var itemsNo = 0;
  
  for (j = 0; j < users.length; j++) { 
    
    //Browser.msgBox (users[j]);
    
    try {
      var page = UrlFetchApp.fetch('https://www.utest.com/api/v1/users/' + users[j]).getContentText();
      var followersNum = page.match(/follower_count":(.*),"following_count/)[1];      
      itemsCount = itemsCount + parseInt(followersNum);
      itemsNo = itemsNo + 1;
    }  
    catch (e)?
    {
      var followersNum = "N/A";
    }
  
    var cell2 = sheetData.getRange(i+1,j+2).setValue(followersNum);
    
  }
  
  //Browser.msgBox (itemsNo + "-" + itemsCount)
  
  var cell3 = sheetData.getRange(i+1,8).setValue(Math.round(itemsCount/itemsNo));
  
  /*
  var cell3 = sheetData.getRange(i+1,8).setValue(Math.round((sheetData.getRange(i+1,2).getValue() +
                                              sheetData.getRange(i+1,3).getValue() +
                                              sheetData.getRange(i+1,4).getValue() +
                                              sheetData.getRange(i+1,5).getValue() +
                                              sheetData.getRange(i+1,6).getValue() +
                                              sheetData.getRange(i+1,7).getValue()
                                             )/itemsNo));
  */
  
  //beware with no data column I and J give error
  var cell4 = sheetData.getRange(i+1,9).setValue(Math.round(sheetData.getRange(i+1,8).getValue()-sheetData.getRange(i,8).getValue()));  
  var cell5 = sheetData.getRange(i+1,10).setValue(((sheetData.getRange(i+1,8).getValue()-sheetData.getRange(i,8).getValue())/sheetData.getRange(i,8).getValue())*100);
  var cell6 = sheetData.getRange(i+1,11).setValue(maxRowValue(sheetData,i+1,2,6));
  
  //alternate background color   
  if (sheetData.getRange('A' + i).getBackground() == '#ffffff') {
    sheetData.getRange(i+1, 1, 1,11).setBackground('#f3f3f3');
  }
  
  //set update date
  sheetData.getRange(4,11).setValue('last update: ' + currentDate)
    
}


//SUPPORT FUNCTIONS

//timezone pad
function pad(number, length){
    var str = "" + number
    while (str.length < length) {
        str = '0'+str
    }
    return str
}

//leading zeros
function addZero(i) {
    if (i < 10) {
        i = "0" + i;
    }
    return i;
}

//find max value
function maxRowValue(sheet, startRow, startColumn, numberCols) {
  
  var min = sheet.getRange(startRow, startColumn).getValue()
  var max = sheet.getRange(startRow, startColumn).getValue()
  
  for (j = 1; j < numberCols+1; j++) { 
       
      //Browser.msgBox('min= ' + min + 'max= ' + max)
   
      if (sheet.getRange(startRow, startColumn + j).getValue() > max) {max = sheet.getRange(startRow, startColumn + j).getValue()}
      if (sheet.getRange(startRow, startColumn + j).getValue() < min) {min = sheet.getRange(startRow, startColumn + j).getValue()}
   
  }
  
  return max-min;
}