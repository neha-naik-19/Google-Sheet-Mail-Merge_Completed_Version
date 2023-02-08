function main() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle("Google Drive File Search")
  );
}

function getSerchedFiles()
{
   var file = DriveApp.getFiles();
   var data = {};
   var e = "";
   data[e] = {};
   data[e].files = [];
   
   while (file.hasNext()) {
      var filedata = file.next();
      data[e].files.push({name: filedata.getName(), id: filedata.getId(), mimeType: filedata.getMimeType(), url: filedata.getUrl()});
      Logger.log(data[e]);
   }
   
   return data;
}

function getLocation(val) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetName = spreadsheet.getSheetName();
  var selectedRange = spreadsheet
    .getSelection()
    .getActiveRange()
    .getA1Notation();
      
  spreadsheet.getRange(selectedRange).setValue(val).setFontColor('#00f'); 

//  spreadsheet.getRange(selectedRange).setWrap(true);
  
  spreadsheet.getRange(selectedRange).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  return {
    sheet: sheetName,
    range: selectedRange,
    value: val
  };
}


