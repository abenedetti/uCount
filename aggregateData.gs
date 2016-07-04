/*
############################
# uCount data aggregator #
############################

This script retrieve data from "data" sheet and aggregate it.

*/

var sheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
var sheetDataAgg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data_agg');


//MAIN FUNCTIONS
//main function
function launchAggregatorEngine() {

  //this is the triggered function that launch the months and years aggregations

  //find the current year and month
  var nowDate = new Date();
  var stringYearDate = nowDate.getFullYear();  
  var stringMonthDate = stringYearDate + addZero(nowDate.getMonth()+1);

  //var stringMonthDate = "201706"; //for testing purposes
  //var stringYearDate = "2017"; //for testing purposes
  
  //months: usually checks only current month
  monthsAggregation(stringMonthDate)
  //set update date
  var cell = sheetDataAgg.getRange(4,3).setValue('last update: ' + nowDate)
  
  //months: usually checks only current month
  yearsAggregation(stringYearDate)  
  //set update date
  var cell = sheetDataAgg.getRange(4,7).setValue('last update: ' + nowDate)
  
}

//months aggregator
function monthsAggregation(stringNow) {
  
  //search the row for the current date   
  var values = sheetDataAgg.getRange("A:A").getValues();  
  var lastRowMonth = lastColumnRow("A", "data_agg")
  
  var matchRow = 5;
    
  while ((matchRow < lastRowMonth) && (values[matchRow] != stringNow)) {matchRow++;}
  matchRow++;
  
  //if now value is not found insert a new row otherwise update existing value
  if ((matchRow-1) == lastRowMonth) {
  
    //Browser.msgBox(stringNow + " not found, insert!");
    monthsCalculation(stringNow, lastRowMonth+1);     
  
  } else {  
    
    //Browser.msgBox(stringNow + " found at row " + matchRow + ", update!");
    monthsCalculation(stringNow, matchRow);
  
  }

}


//year aggregator
function yearsAggregation(stringNow) {
  
  //search the row for the current date   
  var values = sheetDataAgg.getRange("E:E").getValues();  
  var lastRowYear = lastColumnRow("E", "data_agg")
  
  var matchRow = 5;
    
  while ((matchRow < lastRowYear) && (values[matchRow] != stringNow)) {matchRow++;}
  matchRow++;
  
  //if now value is not found insert a new row otherwise update existing value
  if ((matchRow-1) == lastRowYear) {
  
    //Browser.msgBox(stringNow + " not found, insert!");
    yearsCalculation(stringNow, lastRowYear+1);     
  
  } else {  
    
    //Browser.msgBox(stringNow + " found at row " + matchRow + ", update!");
    yearsCalculation(stringNow, matchRow);
  
  }

}

//monthly calculator
function monthsCalculation(monthYear, insertRow) {

  //search in the "data" sheet all the rows with year and month equal to monthYear
  //sum up the variations fields and put the results in the month table of the data_agg sheet
  
  var i = 7;   //because columns variations and variation % are empty for row 6
  var sumVar = 0;
  var sumVarPer = 0;  

  var scanCol = sheetData.getRange(i,1).getValue(); 
  scanCol = scanCol.substring(6, 10) + scanCol.substring(3, 5);
  
  //Browser.msgBox (lastColumnRow("A","data"));
  
  //scan all "data" table
  while (i <= lastColumnRow("A","data")) {
    
    //if values match update the sum
    if (scanCol == monthYear) {    
      sumVar = sumVar + sheetData.getRange(i,9).getValue();
      sumVarPer = sumVarPer + sheetData.getRange(i,10).getValue();    
    }
    
    i = i + 1;    
    scanCol = sheetData.getRange(i,1).getValue(); 
    scanCol = scanCol.substring(6, 10) + scanCol.substring(3, 5);
    
    //Browser.msgBox ("End of round: i= " + i + " sumVar= " + sumVar + " sumVarPer= " + sumVarPer);
    
  }
    
  //populate the months table  
  var cellA = sheetDataAgg.getRange(insertRow,1).setValue(monthYear);
  var cellB = sheetDataAgg.getRange(insertRow,2).setValue(sumVar);
  var cellC = sheetDataAgg.getRange(insertRow,3).setValue(sumVarPer);
  
  //alternate background color   
  if (sheetDataAgg.getRange('A' + (insertRow-1)).getBackground() == '#ffffff') {
    sheetDataAgg.getRange(insertRow, 1, 1, 3).setBackground('#f3f3f3');
  }
}

//yearly calculator
function yearsCalculation(year, insertRow) {

  //search in the "data" sheet all the rows with year year in input argument
  //sum up the variations fields and put the results in the year table of the data_agg sheet
  
  var i = 7;   //because columns variations and variation % are empty for row 6
  var sumVar = 0;
  var sumVarPer = 0;  

  var scanCol = sheetData.getRange(i,1).getValue(); 
  scanCol = scanCol.substring(6, 10);
  
  //Browser.msgBox (lastColumnRow("A","data"));
  
  //scan all "data" table
  while (i <= lastColumnRow("A","data")) {
    
    //if values match update the sum
    if (scanCol == year) {    
      sumVar = sumVar + sheetData.getRange(i,9).getValue();
      sumVarPer = sumVarPer + sheetData.getRange(i,10).getValue();    
    }
    
    i = i + 1;    
    scanCol = sheetData.getRange(i,1).getValue(); 
    scanCol = scanCol.substring(6, 10);
    
    //Browser.msgBox ("End of round: i= " + i + " sumVar= " + sumVar + " sumVarPer= " + sumVarPer);
    
  }
    
  //populate the months table  
  var cellE = sheetDataAgg.getRange(insertRow,5).setValue(year);
  var cellF = sheetDataAgg.getRange(insertRow,6).setValue(sumVar);
  var cellG = sheetDataAgg.getRange(insertRow,7).setValue(sumVarPer);
  
  //alternate background color   
  if (sheetDataAgg.getRange('E' + (insertRow-1)).getBackground() == '#ffffff') {
    sheetDataAgg.getRange(insertRow, 5, 1, 3).setBackground('#f3f3f3');
  }
}


//SUPPORT FUNCTIONS

//last row of specific function
//courtesy of "tinifni"
//original source: http://stackoverflow.com/questions/4169914/selecting-the-last-value-of-a-column
function lastColumnValue(column, sheet) {
  var lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getMaxRows();
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(column + "1:" + column + lastRow).getValues();

  for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
  return values[lastRow - 1];
}


function lastColumnRow(column, sheet) {
  var lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getMaxRows();
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(column + "1:" + column + lastRow).getValues();

  for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
  return lastRow;
}