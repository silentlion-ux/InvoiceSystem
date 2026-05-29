/**
 * This page contains settings used by the rest of the code - things like the names of tabs.
 * 
 * This page also tells the code which sheets (tabs) belong to which artist and which column is where.
 * eg, Which column has the dates of the performances, etc.
 * 
 * Things you're more likely to want to edit are near the top - see notes in green, and refer to handover notes for more information
 */


// ***********************
// Global pseudo-macros
// ***********************
/** Names of the sheets are in red */
const SPECTACULAR_BOOKINGS_SHEET = 'Spectacular Artist';
const SPECTACULAR_INVOICE_SHEET = 'SA Invoices';
const GREAT_BOOKINGS_SHEET = 'Great Musician';
const GREAT_INVOICE_SHEET = 'GM Invoices';
const EXTERNALS_BOOKINGS_SHEET = 'Externals';
const EXTERNALS_INVOICE_SHEET = 'Externals Invoices';
const OTHER_BOOKINGS_SHEET = 'Other';
const OTHER_INVOICE_SHEET = 'Other Invoices';
const SETTINGS_SHEET = 'Settings';
const ADDRESSES = 'Addresses';
const ROW_CELL_ADDRESS = 'O17'; /** <-- where the "Switch 1" number is in the invoice sheets */
const YEAR_CELL = 'Q17'; /** <-- ditto for the year */

// for exporting to Google Drive
const OTHER_FOLDER = 'Other';
// ***********************

// columns to be identified
/** Names of column headings in the booking sheets that the code needs to know */
const cols = ["Month","Artist","Venue","Date","Booked With Venue","Sent Confirmation to Artist","We've received","Due?","Inv Ref","ADS"]

console.log("Collecting settings...");

// ***********************
// Global Declarations
// ***********************
                  /** Leave these alone */
                    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
                    const ui = SpreadsheetApp.getUi();
                    const extBookings = spreadsheet.getSheetByName(EXTERNALS_BOOKINGS_SHEET);
                    const spectacularBookings = spreadsheet.getSheetByName(SPECTACULAR_BOOKINGS_SHEET);
                    const greatBookings = spreadsheet.getSheetByName(GREAT_BOOKINGS_SHEET);
                    const otherBookings = spreadsheet.getSheetByName(OTHER_BOOKINGS_SHEET);
                    const extInvoice = spreadsheet.getSheetByName(EXTERNALS_INVOICE_SHEET);
                    const spectacularInvoice = spreadsheet.getSheetByName(SPECTACULAR_INVOICE_SHEET);
                    const greatInvoice = spreadsheet.getSheetByName(GREAT_INVOICE_SHEET);
                    const otherInvoice = spreadsheet.getSheetByName(OTHER_INVOICE_SHEET);
                    const settings = spreadsheet.getSheetByName(SETTINGS_SHEET);
                    const addresses = spreadsheet.getSheetByName(ADDRESSES);
                    const thisYear = new Date().getFullYear();

                    // for using Google Sheets in-built formulas etc. Date and DateLocale objects in particular have proven less robust than Google Sheets native formulae in this context
                  /** Leave these alone */
                    const inputCell = settings.getRange('O101');
                    const outputCell = settings.getRange('O102');
                    const tertiaryCell = settings.getRange('O103');
                    const colsLogStart = settings.getRange('A110');
                    const finalColNumHolder = colsLogStart;

/** The values in red are cell references, keep these updated */
const emailTemplateSubject = settings.getRange('C37');
const emailTemplateBody1 = settings.getRange('C38');
const emailTemplateBody2 = settings.getRange('C39');
const addressFinder = addresses.getRange('C3');
const addressSpitter = addresses.getRange('F3');
// ***********************

// ***********************
// Settings
// ***********************
/** Keep the values in red updated */
const headFolderId = settings.getRange('O13').getValue().slice(-33);
const sortFiles = settings.getRange('O14').getValue();
const debugFolderId = settings.getRange('O33').getValue().slice(-33);
const useDebugFolder = settings.getRange('O34').getValue();
const fromEmail = settings.getRange('O25').getValue();

/**
 * !! DON'T EDIT BELOW HERE !!
 */
// INDENTIFYING CORRECT COLUMS

// names used in script, don't change 
const cols2 = ["month","artist","venue","date","bookedWithVenue","sentConfirmationToArtist","recStatus","dueStatus","invRef","invGenCol"];
console.log("Sheet and script headings match length? "+(cols.length==cols2.length));

// sheets to be searched
const sheetCols = {
  extObj : {
    "thisSheet":extBookings,
    "thisSheetName":EXTERNALS_BOOKINGS_SHEET,
    "invSheet":extInvoice,
    "headingsRow":6,
    "sheetArtist":"ext"
    },
  spectacularObj : {
    "thisSheet":spectacularBookings,
    "thisSheetName":SPECTACULAR_BOOKINGS_SHEET,
    "invSheet":spectacularInvoice,
    "headingsRow":7,
    "sheetArtist":"Spectacular Artist"
    },
  greatObj : {
    "thisSheet":greatBookings,
    "thisSheetName":GREAT_BOOKINGS_SHEET,
    "invSheet":greatInvoice,
    "headingsRow":7,
    "sheetArtist":"Great Musician"
    },
  otherObj : {
    "thisSheet":otherBookings,
    "thisSheetName":OTHER_BOOKINGS_SHEET,
    "invSheet":otherInvoice,
    "headingsRow":16,
    "sheetArtist":"Other"
    }
}

// aliases for working alongside arrays
Object.defineProperty(sheetCols,"0",{value:sheetCols.extObj});
Object.defineProperty(sheetCols,"1",{value:sheetCols.spectacularObj});
Object.defineProperty(sheetCols,"2",{value:sheetCols.greatObj});
Object.defineProperty(sheetCols,"3",{value:sheetCols.otherObj});

// populating
const finalColNum = finalColNumHolder.getValue();
const colsArray = settings.getRange(colsLogStart.getRow()+1,colsLogStart.getColumn()+1,Object.keys(sheetCols).length,finalColNum).getValues();
let iterator = 0;
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
    let textFinder = sheetCols[x].thisSheet.getRange('A:A').createTextFinder('headings');
    headingsRow = textFinder.findNext().getRow();
    sheetCols[x].headingsRow = headingsRow;
    // for each year within the sheet.
    //NOTE! if you change j here, you must also change '2025' further down
    for (let j=2025; j<2040; j++){
      // checking for the year
      textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j));
      let yearCol = textFinder.findNext();
      if (yearCol != null){
        spreadsheet.toast("Getting columns for "+sheetCols[x].sheetArtist+", "+j);
        // get the range for the year
        yearCol = yearCol.getColumn();
        textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j+1));
        let yearRange = 0;
        let nextYearCol = textFinder.findNext()
        if (nextYearCol != null) {
         nextYearCol = nextYearCol.getColumn();
         yearRange = sheetCols[x].thisSheet.getRange(headingsRow,yearCol,1,nextYearCol-yearCol);
         // writing year and sheet headings to settings sheet
         colsLogStart.offset(0,(j-2025)*cols.length+1).setValue(j);
        }
        else {
          console.log("final year reached");
          spreadsheet.toast("Getting columns for "+sheetCols[x].sheetArtist+", "+j);
          textFinder = sheetCols[x].thisSheet.getRange('1:1').createTextFinder(String(j-1));
          const lastYearCol = textFinder.findNext().getColumn();
          yearRange = sheetCols[x].thisSheet.getRange(headingsRow,yearCol,1,yearCol-lastYearCol);
          colsLogStart.offset(0,(j-2025)*cols.length+1).setValue(j);
          const myKeys = Object.keys(sheetCols);
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
              const endCol = colsLogStart.offset(iterator,((j-2025)*cols.length)+k+1).getColumn();
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

function runEmpty(){}

console.log("Settings collected.");
