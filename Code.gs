function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Data Copy');
  menu.addItem('Copy All', 'copyDataNoFilter');
  menu.addItem('Copy Tagged Only', 'copyDataWithFilter'); //use to transfer data to Client Cleansing file
  menu.addToUi();
}

//trigger functions
function copyDataNoFilter(){
  totalDataCopy_(0); //column A
}

function copyDataWithFilter(){
  totalDataCopy_(16); //tag array column Q = 16
}
//end of trigger functions

/**
 * Main function
 * @param {number} Column to filter out blank cells
 * 0 = no filter.
 * 16 = filter for calendar event tags
 * Transfers data from source sheet to target sheet
 */
let lastSuccessfulRow = 0;

function totalDataCopy_(filterNum, startRow = 0) { 
  //source spreadsheet
  const sourceSheet = getSheet_("B1", "B2");
  //target spreadsheet
  const targetSheet = getSheet_("B4", "B5");

  //get all source data
  const sourceArray = sourceSheet.getDataRange().getValues();
  console.log("sourceArray.length",sourceArray.length)

  //split source data 1000 rows at a time
  let pasteDataArray = []; //save split data here 
  for(let i = startRow; i < sourceArray.length; i++){
    //apply filter; 0 = no filter
    if(sourceArray[i][filterNum] == "") {
      continue;
    } 
    pasteDataArray.push(sourceArray[i]);

    if(pasteDataArray.length == 1000 || i == sourceArray.length-1){ //every 1000 row; change if less than 1000 remain
      try {
        pasteDataToSource_(targetSheet, pasteDataArray, i);
        lastSuccessfulRow = i;
      } catch (error) {
        console.log(error);
        totalDataCopy_(filterNum, lastSuccessfulRow);
      }
      Utilities.sleep(100); // delay for 100 milliseconds
      pasteDataArray = null;
      pasteDataArray = []; //reinitialize pasteDataArray
      }
  }
  //showAlert();
}

/**
 * sub-function to get sheet name
 * @param {string} urlCell Range which has the spreadsheet URL
 * @param {string} sheetNameCell Range which has the sheet name
 * @return Call to sheet
 */
function getSheet_(urlCell, sheetNameCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const thisSheet = ss.getSheetByName('Sheet1');
  try{
    const sheet = thisSheet;
    const spreadsheetURL = sheet.getRange(urlCell).getValue();
    const spreadsheetId = spreadsheetURL.split("/").slice(-2,-1).toString();
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = sheet.getRange(sheetNameCell).getValue();
    const remoteSheet = spreadsheet.getSheetByName(sheetName);

    console.log("id:", spreadsheetId);
    console.log("sheet name:", sheetName);

    return remoteSheet;
  } catch (e){
    console.log(e);
    showAlert();
  }
}

/**
 * sub-function to paste data into source sheet
 * @param {string} Target spreadsheet to paste data into
 */
function pasteDataToSource_(targetSheet, pasteDataArray, i) {
  const last_row = targetSheet.getLastRow();
  const pasteRange = targetSheet.getRange(last_row + 1, 1, pasteDataArray.length, pasteDataArray[0].length);
  pasteRange.setValues(pasteDataArray);
  console.log(i+"行目までのコピペが完了");
  pasteDataArray = null;
}