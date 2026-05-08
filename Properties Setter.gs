/**
 * This page contains settings used by the rest of the code - things like the names of tabs.
 * 
 * This page also tells the code which sheets (tabs) belong to which artist and which column is where.
 * eg, Which column has the dates of the performances, etc.
 */

// ***********************
// Global pseudo-enums
// ***********************
const SPENCE_BOOKINGS_SHEET = 'Spencer Flay';
const SPENCE_INVOICE_SHEET = 'SF Invoices';
const GEORGE_BOOKINGS_SHEET = 'George Croucher';
const GEORGE_INVOICE_SHEET = 'GC Invoices';
const EXTERNALS_BOOKINGS_SHEET = 'Externals';
const EXTERNALS_INVOICE_SHEET = 'Externals Invoices';
const OTHER_BOOKINGS_SHEET = 'Other';
const OTHER_INVOICE_SHEET = 'Other Invoices';
const SETTINGS_SHEET = 'Settings';
const ADDRESSES = 'Addresses';
const ROW_CELL_ADDRESS = 'O17';
const YEAR_CELL = 'Q17';

// for exporting to Google Drive
const OTHER_FOLDER = 'Other';
// ***********************

// columns to be identified
// These are the names used as headings in the bookings sheets, update this if sheet is changed
const cols = ["Month","Artist","Venue","Date","Booked With Venue","Sent Confirmation to Artist","We've received","Due?","Inv Ref","ADS"]

console.log("Collecting settings...");

// ***********************
// Global Declarations
// ***********************
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();
const extBookings = spreadsheet.getSheetByName(EXTERNALS_BOOKINGS_SHEET);
const spenceBookings = spreadsheet.getSheetByName(SPENCE_BOOKINGS_SHEET);
const georgeBookings = spreadsheet.getSheetByName(GEORGE_BOOKINGS_SHEET);
const otherBookings = spreadsheet.getSheetByName(OTHER_BOOKINGS_SHEET);
const extInvoice = spreadsheet.getSheetByName(EXTERNALS_INVOICE_SHEET);
const spenceInvoice = spreadsheet.getSheetByName(SPENCE_INVOICE_SHEET);
const georgeInvoice = spreadsheet.getSheetByName(GEORGE_INVOICE_SHEET);
const otherInvoice = spreadsheet.getSheetByName(OTHER_INVOICE_SHEET);
const settings = spreadsheet.getSheetByName(SETTINGS_SHEET);
const addresses = spreadsheet.getSheetByName(ADDRESSES);
const thisYear = new Date().getFullYear();

// for using Google Sheets in-built formulas etc
const inputCell = settings.getRange('O101');
const outputCell = settings.getRange('O102');
const tertiaryCell = settings.getRange('O103');
const colsLogStart = settings.getRange('A110');
const emailTemplateSubject = settings.getRange('C37');
const emailTemplateBody1 = settings.getRange('C38');
const emailTemplateBody2 = settings.getRange('C39');
const addressFinder = addresses.getRange('C3');
const addressSpitter = addresses.getRange('F3');
const finalColNumHolder = colsLogStart;
// ***********************

// ***********************
// Settings
// ***********************

headFolderId = settings.getRange('O13').getValue().slice(-33);
sortFiles = settings.getRange('O14').getValue();
debugFolderId = settings.getRange('O33').getValue().slice(-33);
useDebugFolder = settings.getRange('O34').getValue();
fromEmail = settings.getRange('O25').getValue();

// INDENTIFYING CORRECT COLUMS

// names used in script, don't change 
const cols2 = ["month","artist","venue","date","bookedWithVenue","sentConfirmationToArtist","recStatus","dueStatus","invRef","invGenCol"];
console.log("Sheet and script headings match? "+(cols.length==cols2.length));

// sheets to be searched
const sheetCols = {
  extObj : {"thisSheet":extBookings,"thisSheetName":EXTERNALS_BOOKINGS_SHEET,"invSheet":extInvoice,"headingsRow":6,"sheetArtist":"ext"},
  spenceObj : {"thisSheet":spenceBookings,"thisSheetName":SPENCE_BOOKINGS_SHEET,"invSheet":spenceInvoice,"headingsRow":7,"sheetArtist":"Spencer Flay"},
  georgeObj : {"thisSheet":georgeBookings,"thisSheetName":GEORGE_BOOKINGS_SHEET,"invSheet":georgeInvoice,"headingsRow":7,"sheetArtist":"George Croucher"},
  otherObj : {"thisSheet":otherBookings,"thisSheetName":OTHER_BOOKINGS_SHEET,"invSheet":otherInvoice,"headingsRow":16,"sheetArtist":"Other"}
}

// aliases for working alongside arrays
Object.defineProperty(sheetCols,"0",{value:sheetCols.extObj});
Object.defineProperty(sheetCols,"1",{value:sheetCols.spenceObj});
Object.defineProperty(sheetCols,"2",{value:sheetCols.georgeObj});
Object.defineProperty(sheetCols,"3",{value:sheetCols.otherObj});

// populating
finalColNum = colsLogStart.getValue();
colsArray = settings.getRange(colsLogStart.getRow()+1,colsLogStart.getColumn()+1,Object.keys(sheetCols).length,finalColNum).getValues();
var iterator = 0;
for (let x in sheetCols){
  for (let i=2025; i<2040; i++) {
    sheetCols[x][String(i)] = {};
    for (let j=0; j<cols.length; j++){
      sheetCols[x][String(i)][cols2[j]]=colsArray[iterator][(i-2025)*cols.length+j];
      //console.log(cols[j]+" in "+i+" in "+sheetCols[x].thisSheetName+" is at column "+sheetCols[x][i][cols2[j]]);
    }
    sheetCols[x][i].leftCol = Math.min( ...colsArray[iterator].slice((i-2025)*cols2.length,(i-2025)*cols2.length+cols2.length));
  }
  iterator++;
}

function resetCols(){
  // clear previous record
  settings.getRange(colsLogStart.getRow(),colsLogStart.getColumn(),Object.keys(sheetCols).length+1,(Object.keys(sheetCols[1]).length*cols2.length)).clear();

  // for each bookings sheet
  iterator = 0;
  for (let x in sheetCols){
    // get the headings row
    iterator++;
    colsLogStart.offset(iterator,0).setValue(sheetCols[x].thisSheetName);
    textFinder = sheetCols[x].thisSheet.getRange('A:A').createTextFinder('headings');
    headingsRow = textFinder.findNext().getRow();
    sheetCols[x].headingsRow = headingsRow;
    // for each year within the sheet.
    //NOTE! if you change j here, you must also change '2025' further down
    for (let j=2025; j<2040; j++){
      // check for the year
      textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j));
      yearCol = textFinder.findNext();
      if (yearCol != null){
        // get the range for the year
        yearCol = yearCol.getColumn();
        textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j+1));
        let yearRange = 0;
        nextYearCol = textFinder.findNext()
        if (nextYearCol != null) {
         nextYearCol = nextYearCol.getColumn();
         yearRange = sheetCols[x].thisSheet.getRange(headingsRow,yearCol,1,nextYearCol-yearCol);
         // writing year and sheet headings to settings sheet
         colsLogStart.offset(0,(j-2025)*cols.length+1).setValue(j);
        }
        else {
          console.log("final year reached");
          textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j-1));
          let lastYearCol = textFinder.findNext().getColumn();
          yearRange = sheetCols[x].thisSheet.getRange(headingsRow,yearCol,1,yearCol-lastYearCol);
          colsLogStart.offset(0,(j-2025)*cols.length+1).setValue(j);
          myKeys = Object.keys(sheetCols);
          if (myKeys.indexOf(iterator)==myKeys.length-1){
            console.log("final sheet in final year");
            colsLogStart.setValue((j-2024)*cols.length+1);
          }
        }
        // identify each column
        for (let k=0; k<cols2.length; k++){
          textFinder = yearRange.createTextFinder(cols[k]);
          let thisCol = textFinder.findNext();
          if (thisCol != null) {
            sheetCols[x][j][cols2[k]]=thisCol;
            console.log(cols[k]+" in "+j+" in "+sheetCols[x].thisSheetName+" is at column "+sheetCols[x][String(j)][cols2[k]].getColumn());
            // writing result to sheet
            colsLogStart.offset(iterator,((j-2025)*cols.length)+k+1).setValue(thisCol.getColumn());
            if (k+1==cols2.length){
              endCol = colsLogStart.offset(iterator,((j-2025)*cols.length)+k+1).getColumn();
              if (endCol > colsLogStart.getValue()){
                colsLogStart.setValue(endCol);
              }
            }
          }
        }
      }
    }     
  }
}

//cleanup
inputCell.clear;
outputCell.clear;

function runEmpty(){}

console.log("Settings collected.");
