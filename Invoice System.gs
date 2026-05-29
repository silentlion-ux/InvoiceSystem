/** @OnlyCurrentDoc */

/** 
 *****FILE MAP*****

This section describes the functions that come below.
 <--- See also the Properties Setter.gs file

* getInvoiceRefs() { fills in invoice refs for a given bookings sheet }

* folderFinder() { finds the correct folder to file the invoice into }

* createPDF() { creates an invoice PDF and names it using the setPDFName() function }

* setPDFName() { creates the name to use as the exported PDF name }

* runExporter() { 
  grabs the user's input settings;
  runs createPDF() with the user's settings and exports to a Google Folder using folderFinder();
  generates download url;
}

These functions run when you press the "export" button on the invoice tabs:
* externalsExportButton() { calls runExporter() for the current invoice in the Externals Invoices tab }
* spectacularExportButton() { ditto for Spectacular Artist }
* greatExportButton() { ditto for Great Musician }
* otherExportButton() { ditto for Other }

* emailer() { drafts an email with an invoice PDF attached }

* exportIterator() { runs through a bunch of invoices, calling runExporter() for each }

* getFullRange() { gets the "invoice due" row that's furthest down in a given booking sheet across multiple years, used to help other functions }

* emailFactory() { similer to exportIterator(), but also runs emailer() for each invoice }

This function runs when you press the "RUN" button on the Settings tab:
* factory() { runs all of the above functions, depending on user settings in the Settings tab }

* onEdit() { if the user changes the "Switch 1" number on any of the invoices tabs, it removes the download link }

* sillyDebug() { for testing code snippets }

*****************************************
*/

// fills in invoice refs for due invoice rows
// the obj parameter contains information about the artist and their sheet. It's created in Properties Setter.gs
function getInvoiceRefs(obj, year=thisYear) {

  console.log("setting invoice refs for "+obj.thisSheetName+", "+year);
  spreadsheet.toast("Filling invoice refs for "+obj.thisSheetName," "+year);

  const dueRows = getFullRange(obj, year)[year];
  if (dueRows[0]>0){
    const confs = obj.thisSheet.getRange(1,obj[year].bookedWithVenue,dueRows.at(-1),1).getValues();
    for (let i=0; i<dueRows.length; i++){
      if (confs[dueRows[i]-1][0]=="Confirmed"){
        console.log("filling ref at row "+dueRows[i]);
        obj.thisSheet.getRange(obj.headingsRow,obj[year].invGenCol).copyTo(
        obj.thisSheet.getRange(dueRows[i],obj[year].invRef),
        SpreadsheetApp.CopyPasteType.PASTE_VALUES,
       false,);
       // waiting for the ref template cell to generate new reference number
       while (obj.thisSheet.getRange(6,obj[year].invGenCol).getValue()==obj.thisSheet.getRange(dueRows[i],obj[year].invRef).getValue()){}
      }
    }
  }
}

// finds the correct Google Drive folder to export to. If it doesn't exist, offers to export to the main Invoices folder
function folderFinder (idOfHeadFolder, artist="who", year=thisYear) {
  const headFolder = DriveApp.getFolderById(idOfHeadFolder);
  let folder = headFolder.getFoldersByName(year);
  if (!folder.hasNext()){
    console.log(year+" folder not found. Waiting for user response.");
    let response = ui.alert("ERROR","Google Drive folder not found for " +artist+ " " +year+ ". Export to invoices folder?", ui.ButtonSet.YES_NO,);
      if (response === ui.Button.NO) { spreadsheet.toast("Export cancelled."); return null; }
      if (response === ui.Button.YES) { return headFolder; }
  }
  if (artist == "ext"){
    folder = folder.next().getFoldersByName(OTHER_FOLDER);
    return folder.next();
  }
  try {
    folder = folder.next().getFoldersByName(artist);
    return folder.next();
  } catch(err) {
    console.log(year+" "+artist+" folder not found. Waiting for user response.");
    let response = ui.alert("ERROR","Google Drive folder not found for " +artist+ " " +year+ ". Export to invoices folder?", ui.ButtonSet.YES_NO,);
      if (response === ui.Button.NO) { spreadsheet.toast("Export cancelled."); return null; }
      if (response === ui.Button.YES) { return headFolder; }
  }
}

/**
 * exports a PDF to a chosen GDrive Folder, returning the PDF from GDrive
 * parameters (the bits in the (parentheses)):
 * ssId - Id of the Google Spreadsheet
 * sheet - Sheet to be converted as PDF
 * pdfName - File name of the PDF being created
 * targFolder - ID of the GDrive folder to export to
 * returns the PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName, targFolder) {
  const fr = 4, fc = 2, lc = 11, lr = 46;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?ExportFormat=pdf&format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", muteHttpExceptions: true, headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');
  pdfFile = targFolder.createFile(blob);
  return pdfFile;
}

// Sets the PDF name, used inside createPDF():
function setPDFName (obj, artist) {
  let inv = obj.invSheet.getRange(ROW_CELL_ADDRESS).getValue();
  let inYear = obj.invSheet.getRange(YEAR_CELL).getValue();
  let venue = obj.invSheet.getRange('I18').getValue();
  if (venue=="#N/A"){ venue = "Performance"; }
  // if an external artist, get the artist name from the bookings sheet
  if (artist=="ext") { artist = obj.thisSheet.getRange(inv,obj[inYear].artist).getValue(); }

  // figure out the date. JS Date and datelocale objects have proven less robust than using Sheets formulae in this context
  inputCell.setValue(`=4+MATCH(` + inYear + `,'` + obj.thisSheetName + `'!1:1,0)`);
  d = obj.thisSheet.getRange(inv,obj[inYear].date).getValue();
  if (typeof(d)=="string"){
    var yearString = d.slice(-4);
    var monthString = d.slice(-7,-5);
    var dayString = d.slice(0,2);
    d = new Date(yearString+"-"+monthString+"-"+dayString);
  }

  // compensating for daylight savings
  inputCell.setValue(`=indirect("'` + obj.thisSheet.getSheetName() + `'!R"&'` + obj.invSheet.getSheetName() + `'!`+ROW_CELL_ADDRESS+`&"C`+obj[inYear].date+`",0)`);
  outputCell.setValue(`=AND(datevalue(`+inputCell.getA1Notation()+`)>datevalue("29/3/`+inYear+`"),datevalue(`+inputCell.getA1Notation()+`)<datevalue("27/10/`+inYear+`"))`);
  if (outputCell.getValue()==1) {d.setMinutes(d.getMinutes() + 60);}
  let date = d.getUTCDate() + "-" + (d.getUTCMonth()+1) + "-" + d.getUTCFullYear();

  // cleanup
  inputCell.clear();
  outputCell.clear();

  console.log("pdfName: "+ venue + " - " + artist + " " + date);

  return venue + " - " + artist + " " + date;
}

// prepares an invoice and then runs createPDF(), then generates an url if appropriate
function runExporter(obj, artist, download = false) {

  //checking for ref and confirmation
  let year = obj.invSheet.getRange(YEAR_CELL).getValue();
  let refCol = obj[year].invRef;
  let confCol = obj[year].bookedWithVenue;
  let invRow = obj.invSheet.getRange(ROW_CELL_ADDRESS).getValue();
  if (obj.thisSheet.getRange(invRow,refCol).isBlank() || obj.thisSheet.getRange(invRow,confCol).getValue()!="Confirmed") {
    console.log("Booking is missing ref and/or not confirmed, waiting on user input");
    let response = ui.alert("Missing Information",artist+" Booking at row " + invRow + " is not confirmed and/or missing ref number. Are you sure you want to export this invoice?", ui.ButtonSet.YES_NO,);
          if (response === ui.Button.NO) { return; }
  };

  //fetching Gdrive output folder
  if (sortFiles) {
    console.log("sorting");

    if (useDebugFolder){
      console.log("using debug folder");
      folder = DriveApp.getFolderById(debugFolderId);
    }

    if (!useDebugFolder) {
      folder = folderFinder(headFolderId, artist, year);
    }
  }

  // getting artist from Other if necessary
  if (artist=="Other"){
    artist = obj.thisSheet.getRange(invRow,obj[year].artist).getValue();
    spreadsheet.toast("Artist: "+artist);
  }

  if (folder == null){ return null; }

  // exporting PDF
  let pdfFile = createPDF(spreadsheet.getId(),obj.invSheet,setPDFName(obj, artist),folder);
  console.log("PDF exported");

  if (download==true) {
    let downloader = pdfFile.getDownloadUrl();
    obj.invSheet.getRange('N2').setValue('=hyperlink("'+ downloader +'","DOWNLOAD")');
  }
  return pdfFile;
}


// button functions that run when you press one of the "export" buttons
function externalsExportButton(download = true) {
  return runExporter(sheetCols.extObj, "ext", download);
}

function spectacularExportButton(download = true) {
  return runExporter(sheetCols.spectacularObj, "Spectacular Artist", download);
}

function greatExportButton(download = true) {
  return runExporter(sheetCols.greatObj, "Great Musician", download);
}

function otherExportButton(download = true) {
  return runExporter(sheetCols.otherObj, "Other", download);
}

// creating an individual gmail draft
function emailer(address="hello@example.com",info=[], attachment) {
  // applying override email address if indicated in user settings
  if (settings.getRange('O23').getValue()){address = settings.getRange('O22').getValue();}
  
  // getting the "from" email address, or using user's if n/a
  let fromEmailAddress = String(fromEmail);
  if (fromEmailAddress = "" || !fromEmailAddress.includes("@")){ fromEmailAddress = Session.getActiveUser().getEmail();}

  infoString = "";
  for (let i=0; i<info.length; i++){
    for (let j=0; j<info[i].length-1; j++){
      infoString = infoString.concat(info[i][j]+", ");
    }
    infoString = infoString.concat(info[i][info[i].length-1]+"\n");
  }
  console.log("email info: "+infoString);
  GmailApp.createDraft(
    address,
    emailTemplateSubject.getValue(),
    emailTemplateBody1.getValue()+"\n"+infoString+emailTemplateBody2.getValue(),
    {
      attachments: attachment.map(blb => blb.getAs(MimeType.PDF)),
      name: 'Alex Dale-Staples',
      from: fromEmailAddress,
    },
  );
}

// exports all selected invoices on a given sheet
// the "obj" parameter contains info about the artist and their sheet, created in Properties Setter.gs
// the "exporter" parameter points to an export button function, eg "spectacularExportButton()"
function exportIterator(obj=sheetCols.spectacularObj, exporter) {
  let last = 300;
  let venueCheck = settings.getRange('H23').getValue();
  let rowCell = obj.invSheet.getRange(ROW_CELL_ADDRESS);
  let holdOldRow = rowCell.getValue();
  for (let j=0; j<2; j++){
    if (obj[thisYear+j].dueStatus){
      recStats = obj.thisSheet.getRange(1,obj[thisYear+j].dueStatus,last,1).getValues();
      venues = obj.thisSheet.getRange(1,obj[thisYear+j].venue,last,1).getValues();
      for (let i=0; i<last; i++) {
        if (recStats[i][0]=="DUE: Invoice" && venues[i][0].includes(venueCheck)){    
          rowCell.setValue(i+1);
          console.log("Exporting " + obj.thisSheetName + " row " + rowCell.getValue());
          spreadsheet.toast("Exporting "+ obj.sheetArtist + " row " + rowCell.getValue());
          pdfFile = exporter();
        }
      }
    }
  }

  // cleanup
  rowCell.setValue(holdOldRow);
}

// gets the Invoice: DUE row that is furthest down for a given sheet, across multiple years
function getFullRange(obj, ...years) {
  let dueArray = [[],[]];
  let bigLastDue = 0; // this will hold the furthest down row
  for (let year of years){
    dueArray[year]=[];
    const dueColNum = obj[year].dueStatus;
    const dueColLtr = obj.thisSheet.getRange(1,dueColNum).getA1Notation().charAt(0);
    const dueColVals = obj.thisSheet.getRange(dueColLtr+":"+dueColLtr).getValues();
    let lastDue = 0;
    for (let i = 0; i<dueColVals.length; i++){
        lastDue = dueColVals[i][0] == "DUE: Invoice" ? i + 1 : lastDue;
        if (dueColVals[i][0] == "DUE: Invoice"){
          dueArray[year].push(i+1);
        }
    }
    bigLastDue = lastDue > bigLastDue ? lastDue : bigLastDue;
  }
  return dueArray;
}

// exports PDFs as selected by the user in Settings tab and attaches them to draft emails
// It's a bit of a beast, but what it does:
// 1. It checks each booking sheet and year that's fed to it for due bookings.
// 2. For each due booking it finds, it checks each booking sheet and year again for any other bookings that belong to the same venue
// 3. It exports all of the invoices from step 2. and collects them together
// 4. It drafts an email to the venue with all of their invoices from 3. attached.
function emailFactory (ranges, ...years) {
  var venueCheck = settings.getRange('H23').getValue(); // fetching the venue (if any) the user specified in the settings tab
  
  for (let i=0; i<ranges.length; i++) {
    if (ranges[i]!=undefined){ // looking for due bookings in each tab (sheet)...
      for (let year of years){ // ... in each year within that tab (sheet)...
        for (let j=0; j<ranges[i].length; j++) { // ... and now going through each row within that year
          let ven = ranges[i][j][sheetCols[i][year].venue-1]; // fetching the venue for the current row and storing it in my "ven" variable
          if (ranges[i][j][sheetCols[i][year].dueStatus-1]=="DUE: Invoice"&&ven.includes(venueCheck)){ // checking if the booking is marked as due, and that the venue is the one specified by the user in the Settings sheet (if any)
            console.log(Object.keys(sheetCols)[i]+" "+[j+1]+" is due");
            infoArray = []; // this will collect info about each booking to export
            attachArray = []; // this will collect the exported PDFs
            // now we're going to check each sheet, year and booking again for any due invoices at the same venue, so we can attach them together into the same email
            for (let k=0; k<ranges.length; k++){ // for each sheet
              if (ranges[k]!=undefined){ // this stops a lot of errors
                for (let year of years){ // for each year
                  for (let l=0; l<ranges[k].length; l++){ // for each row
                    if (ranges[k][l][sheetCols[k][year].dueStatus-1]=="DUE: Invoice"&&ranges[k][l][sheetCols[k][year].venue-1]==ven){ // again if it's due etc
                      artist = ranges[k][l][sheetCols[k][year].artist-1]; // grabbing the artist name
                      if (artist==undefined){ artist = sheetCols[k].sheetArtist; }
                      let dateVal = ranges[k][l][sheetCols[k][year].date-1]; // grabbing the date
                      // compensating for daylight savings
                      inputCell.setValue(dateVal);
                      outputCell.setValue(`=AND(datevalue(`+inputCell.getA1Notation()+`)>datevalue("29/3/`+year+`"),datevalue(`+inputCell.getA1Notation()+`)<datevalue("27/10/`+year+`"))`);
                      if (outputCell.getValue()==1) {dateVal.setMinutes(dateVal.getMinutes() + 60);}
                      // turning the date into a presentable string
                      dateVal = dateVal.getUTCDate() + "-" + (dateVal.getUTCMonth()+1) + "-" + dateVal.getUTCFullYear();

                      // setting up an exporting the invoice for the current booking
                      infoArray.push([artist,dateVal,ven]);
                      sheetCols[k].invSheet.getRange(ROW_CELL_ADDRESS).setValue(l+1);
                      sheetCols[k].invSheet.getRange(YEAR_CELL).setValue(year);
                      spreadsheet.toast("Exporting "+sheetCols[k].thisSheetName+" row "+[l+1]);
                      attachArray.push(runExporter(sheetCols[k], sheetCols[k].sheetArtist)); // exporting the invoice and storing it in our attachArray[] we made earlier

                      // if the correct Google Drive folder doesn't exist and the user doesn't want to export, removing from collection:
                      if (attachArray.at(-1) == null){
                        console.log("removing element");
                        attachArray.pop();
                        infoArray.pop();
                      }
                      console.log("attached: " + attachArray);
                      ranges[k][l][sheetCols[k][year].dueStatus-1]="INV ATTACHED"; // changing the entry so we don't export it again
                    }
                  }
                }
              }
            }
            // fetching the venue's email address
            spreadsheet.toast("Drafting email to "+ven);
            addressFinder.setValue(ven);
            let venEmail = "";
            let addr = addressSpitter.getValue();
            if (addr!="#N/A"){venEmail = addresses.getRange(addr).offset(-1,0).getValue();}
            console.log("Email to: "+venEmail);
            emailer(venEmail, infoArray, attachArray); // drafting the email with all the invoices attached that belong to the same venue
          }
        }
      }
    }
  }
}

// this runs when you press "RUN" in the Settings tab:
function factory() {
  spreadsheet.toast("Collecting settings");
  ranges = [];
  for (let i=0; i<Object.keys(sheetCols).length; i++) { 
    lastCol = sheetCols[i][thisYear+1].invGenCol-1;
    lastRow = getFullRange(sheetCols[i],thisYear,thisYear+1).at(-2);
    console.log(sheetCols[i].sheetArtist + " invGenColumn: "+sheetCols[i][thisYear+1].invGenCol+" cols: "+lastCol+" rows: "+lastRow);
    if (lastRow>0){ ranges[i] = sheetCols[i].thisSheet.getRange(1,1,lastRow,lastCol).getValues(); }
  }

  // generating invoice refs
  // if user selected "all" in Settings tab:
  if (settings.getRange('E6').getValue()) {
    for (i=thisYear; i<thisYear+2; i++){
      getInvoiceRefs(sheetCols.extObj, i);
      getInvoiceRefs(sheetCols.spectacularObj, i);
      getInvoiceRefs(sheetCols.greatObj, i);
      getInvoiceRefs(sheetCols.otherObj, i);
    }
  }

  // generating invoice refs
  // whichever the user selected in Settings tab (if not "all"):
  else {
    for (i=thisYear; i<thisYear+2; i++){
      //generate for externals
      if (settings.getRange('E7').getValue()) { getInvoiceRefs(sheetCols.extObj, i); }
      // generate for Spectacular
      if (settings.getRange('E8').getValue()) { getInvoiceRefs(sheetCols.spectacularObj, i); }
      // generate for Great
      if (settings.getRange('E9').getValue()) { getInvoiceRefs(sheetCols.greatObj, i); }
      // generate for Other
      if (settings.getRange('E10').getValue()) { getInvoiceRefs(sheetCols.otherObj, i); }
    }
  }

  if (settings.getRange('E21').getValue()){
    emailFactory(ranges,thisYear,thisYear+1);
  }
  
  else {
    // PDF exports
    // getting selected artists from settings tab
    extExpFlag = settings.getRange('E14').getValue() || settings.getRange('E15').getValue();
    spectacularExpFlag = settings.getRange('E14').getValue() || settings.getRange('E16').getValue();
    greatExpFlag = settings.getRange('E14').getValue() || settings.getRange('E17').getValue();
    otherExpFlag = settings.getRange('E14').getValue() || settings.getRange('E18').getValue();

      // exporting from Externals tab
      if (extExpFlag){ exportIterator(sheetCols.extObj, externalsExportButton); }

      // exporting from Spectacular Artist tab
      if (spectacularExpFlag){ exportIterator(sheetCols.spectacularObj, spectacularExportButton); }

      // exporting from Great Musician tab
      if (greatExpFlag){ exportIterator(sheetCols.greatObj, greatExportButton) };

      // exporting from Other tab
      if (otherExpFlag){ exportIterator(sheetCols.otherObj, otherExportButton) };
  }
}

// resets PDF download url
function onEdit(e) {
  sheet = spreadsheet.getActiveSheet();
  shName = sheet.getSheetName();
  if (sheet.getActiveCell().getA1Notation()==ROW_CELL_ADDRESS && (shName==EXTERNALS_INVOICE_SHEET || shName==SPECTACULAR_INVOICE_SHEET || shName==GREAT_INVOICE_SHEET || shName==OTHER_INVOICE_SHEET)) {
    sheet.getRange('N2').clear();
  }
}

// for ad hoc testing
function sillyDebug() {
  //
}
