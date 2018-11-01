//https://medium.com/@aio.phnompenh/make-ocr-tool-in-google-spreadsheet-to-extract-text-from-image-or-pdf-using-google-app-script-c478d4062b8c
//*******************************************************************************************************************//*******************************************************************************************************************
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OCR Tools')
  .addItem('Extract A Single Highlighted Cell', 'doOCR')
  .addItem('Extract Multiple Highlighted Cells', 'doOCRALL')
  .addToUi();
}
//*******************************************************************************************************************

//*******************************************************************************************************************
//Perform OCR on all items in the column

function AutodoOCR() {
  
  //Force focus on  'OCR (Computers)' sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var computers = ss.getSheetByName('OCR (Computers)');
  removeEmptyRows();
  var lastRow = computers.getLastRow();
  Logger.log("lastRow:" + lastRow);
  //This is the column with the URL to the image
  var activeCol = 5;
  //var selected = computers.getRange(1, activeCol, lastRow, 1).getValues().length;
  //Logger.log('selected length is '+selected);
  var today = new Date();
  Logger.log('Date is '+today);
  var archiveFolder = DriveApp.getFolderById('15-B08m_5EUpVdqk65VuiAJq1g4bu09Lk');
  //var selected = SpreadsheetApp.getActiveSheet().getActiveRange().getValues().length;
  //for (var i = 0; i < selected; i++) {
  
  //var activeRow = 1 + i;
  
  if (computers.getRange(lastRow, activeCol + 1).isBlank()){
    Logger.log('AutoOCR performed, Cell Content is '+computers.getRange(lastRow, activeCol + 1).getDisplayValue());
    var valueURL = computers.getRange(lastRow , activeCol).getValue();
    Logger.log('valueURL is '+valueURL);
    if (valueURL != '#REF!'){
      
      try{
        
        var image = UrlFetchApp.fetch(valueURL).getBlob();
        
        
        var file = {
          title: 'OCR File '+today,
          mimeType: 'image/png'
        };
        
        // OCR is supported for PDF and image formats
        file = Drive.Files.insert(file, image, {ocr: true});
        var doc = DocumentApp.openByUrl(file.embedLink);
        var body = doc.getBody().getText();
        //Get link Doc that Generated
        computers.getRange(lastRow, activeCol + 2).setValue(file.embedLink);
        //Get Content of Doc that Generated
        computers.getRange(lastRow, activeCol + 1).setValue(body);
        
        //Move OCR file from root Google Drive to picture folder
        
        Logger.log('file.id is ' + file.id);
        var fileByID = DriveApp.getFileById(file.id);
        archiveFolder.addFile(fileByID);
        DriveApp.getRootFolder().removeFile(fileByID);
        
        //Backup picture file to folder
        /* Logger.log(archiveFolder.getName());
        var add = archiveFolder.createFile(image);
        var attName = 'OCR File '+today;
        add.setName(attName);
        Logger.log(archiveFolder.getFiles());*/
        
      } catch(e) {
        Logger.log('Invalid URL');
      }
      
      //throw("testing complete");
      //MailApp.sendEmail({to:'rmccal14+logger@uncc.edu',subject: "OCR Log!",body: Logger.getLog()});
    }
  } else {
    Logger.log('Cell Content is '+computers.getRange(lastRow, activeCol + 1).getDisplayValue());
    Logger.log('AutoOCR not performed');
  }
  //}
}
//*******************************************************************************************************************
function doOCR(image) {
  //
  var activeCol = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
  var activeRow = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  
  var activeCol2 = SpreadsheetApp.getActiveSheet().getDataRange().getLastColumn()
  var activeRow2 = SpreadsheetApp.getActiveSheet().getDataRange().getLastColumn()
  
  var valueURL = SpreadsheetApp.getActiveSheet().getRange(activeRow, activeCol).getValue();
  
  var image = UrlFetchApp.fetch(valueURL).getBlob();
  
  var file = {
    title: 'OCR File',
    mimeType: 'image/png'
  };
  
  // OCR is supported for PDF and image formats
  file = Drive.Files.insert(file, image, {ocr: true});
  var doc = DocumentApp.openByUrl(file.embedLink);
  var body = doc.getBody().getText();
  
  
  // Print the Google Document URL in the console
  Logger.log("body: %s", body);
  Logger.log("File URL: %s", file.embedLink);
  //Get link Doc that Generated
  SpreadsheetApp.getActiveSheet().getRange(activeRow, activeCol + 2).setValue(file.embedLink);
  //Get Content of Doc that Generated
  SpreadsheetApp.getActiveSheet().getRange(activeRow, activeCol + 1).setValue(body);
}
//*******************************************************************************************************************
function doOCRALL() {
  var selected = SpreadsheetApp.getActiveSheet().getActiveRange().getValues().length;
  for (var i = 0; i < selected; i++) {
    var activeCol = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
    var activeRow = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
    var valueURL = SpreadsheetApp.getActiveSheet().getRange(activeRow + i, activeCol).getValue();
    
    var image = UrlFetchApp.fetch(valueURL).getBlob();
    
    var file = {
      title: 'OCR File',
      mimeType: 'image/png'
    };
    
    // OCR is supported for PDF and image formats
    file = Drive.Files.insert(file, image, {ocr: true});
    var doc = DocumentApp.openByUrl(file.embedLink);
    var body = doc.getBody().getText();
    //Get link Doc that Generated
    SpreadsheetApp.getActiveSheet().getRange(activeRow + i, activeCol + 2).setValue(file.embedLink);
    //Get Content of Doc that Generated
    SpreadsheetApp.getActiveSheet().getRange(activeRow + i, activeCol + 1).setValue(body);
    
  }
}
//*******************************************************************************************************************
//https://stackoverflow.com/questions/44579300/how-to-ignore-empty-cell-values-for-getrange-getvalues
//Delete empty rows
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getSheetByName("OCR (Computers)");
    var maxRows = sheet.getMaxRows(); 
  var result = sheet.getRange("A1:A").getValues();
  var lastRow = [i for each (i in result) if (isNaN(i))].length;
    Logger.log("sheet is "+sheet.getName());
  Logger.log("lastRow is "+lastRow);
    if (maxRows-lastRow > 1){
      Logger.log("Delete Rows");
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    } else {
      Logger.log("Don't Delete Rows");
    }
}
//*******************************************************************************************************************