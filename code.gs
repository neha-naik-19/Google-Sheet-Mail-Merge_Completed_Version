function onInstall(e)
{
  onOpen(e);
}

/* What should the add-on do when a document is opened */
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or SlidesApp or FormApp.
  
  menu.addItem("Mail Merge", "doGet");
  menu.addItem("Google Drive Side Bar", "main");
  menu.addToUi();
}

//convert to propercase
function properCase(phrase) {
  var regFirstLetter = /\b(\w)/g;
  var regOtherLetters = /\B(\w)/g;
  function capitalize(firstLetters) {
    return firstLetters.toUpperCase();
  }
  function lowercase(otherLetters) {
    return otherLetters.toLowerCase();
  }
  var capitalized = phrase.replace(regFirstLetter, capitalize);
  var proper = capitalized.replace(regOtherLetters, lowercase);

  return proper;
}

//get id from google drive url
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

var notvalid = 0;

function doGet()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // First row of data to process
  var numcolumns = 29+1; // Number of rows to process
  
  const lr = sheet.getLastRow()-1;
  
  if(lr == 0)
  {
    Browser.msgBox('No data found in sheet.\\n Kindly Check.');
    
    return
  }
  
  //Check empty cell for Greeting column
  var columnGreeting = sheet.getRange(startRow, 2, lr).getValues();
  var cellrange = sheet.getRange(startRow, 2, lr).getA1Notation();
  
  for(var i=0; i<columnGreeting.length ; i++){
    if (columnGreeting[i] ==""){
      Browser.msgBox('Greeting is empty.\\n Kindly Check. cell: '+ cellrange);
      notvalid = 1;
      break;
    }
  }
    
  if(notvalid == 1)
  {
    return;
  }
  
  //Check empty cell for Receiver Email Address
  var columnReceivermail = sheet.getRange(startRow, 3, lr).getValues();
  var cellrange = sheet.getRange(startRow, 3, lr).getA1Notation();
  
  for(var i=0; i<columnReceivermail.length ; i++){
    if (columnReceivermail[i] ==""){
      Browser.msgBox('Receiver Email Address is empty.\\n Kindly Check. cell: '+ cellrange);
      notvalid = 1;
      break;
    }
  }
  
  if(notvalid == 1)
  {
    return;
  }
  
  //Check empty cell for Subject
  var columnSubject = sheet.getRange(startRow, 4, lr).getValues();
  var cellrange = sheet.getRange(startRow, 4, lr).getA1Notation();
  
  for(var i=0; i<columnSubject.length ; i++){
    if (columnSubject[i] ==""){
      Browser.msgBox('Subject is empty.\\n Kindly Check. cell: '+ cellrange);
      notvalid = 1;
      break;
    }
  }
  
  if(notvalid == 1)
  {
    return;
  }
  
  //Check Mandatory message body
  var columnMessagebody = sheet.getRange(startRow, 5, lr).getValues();
  var cellrange = sheet.getRange(startRow, 5, lr).getA1Notation();
  
  for(var i=0; i<columnMessagebody.length ; i++){
    if (columnMessagebody[i] ==""){
      Browser.msgBox('Mandatory message body is empty.\\n Kindly Check. cell: '+ cellrange);
      notvalid = 1;
      break;
    }
  }
  
  if(notvalid == 1)
  {
    return;
  }
  
  // Fetch the range of cells Exa:A2:B3
  var dataRange = sheet.getRange(startRow, 1, lr, numcolumns);
  var data = dataRange.getDisplayValues();
  
  // Fetch values for each row in the Range.
  for (var i in data) {
    var row = data[i];
    
    var greeting = 'Dear ' + properCase(row[1].trim()) + ',';

    var messagebody = '';
    var msg_arr = [];
    var para1 = '';
    var para2 = '';
    var para3 = '';
    var para4 = '';

    var br = '<br>';
    
    msg_arr[0]= row[5].trim();
    msg_arr[1]= row[6].trim();
    msg_arr[2]= row[7].trim();
    msg_arr[3]= row[8].trim();
    msg_arr[4]= row[9].trim();
    msg_arr[5]= row[10].trim();

    if(row[11].length > 0)
    {
      para1 = row[11].trim();
    }

    if(row[12].length > 0)
    {
      para2 = row[12].trim();
    }

    if(row[13].length > 0)
    {
      para3 = row[13].trim();
    }

    if(row[14].length > 0)
    {
      para4 = row[14].trim();
    }
    
    if(row[5] == '' && row[6] == '' && row[7] == '' && row[8] == '' && row[9] == '' && row[10] == '')
    {
      messagebody = row[4].trim();
    }
    else
    {
      messagebody = row[4].trim();
    
      for (var i = 0; i < msg_arr.length; i++)
      {
        if(msg_arr[i] != '')
        {
          messagebody = messagebody + " " + msg_arr[i];
        }
      }
    }

    //    var htmlmessage = 
    //    '<body>' + 
    //    '<h4>' + greeting + '</h4>' +
    //    '<p> \n' + messagebody +
    //  '<p> with regards,<br>Department of Computer Science & Information Systems (CS&IS)<br>Birla Institute of  Technology   & Science (BITS), Pilani<br>K K Birla Goa Campus<br>NH 17B, Zuarinagar<br>Goa, India. 403 726<br>Phone:     +91 0832 2580851<br>E Mail: csis.office@goa.bits-pilani.ac.in </p>' +
    //    '</body>'

    var recipient = row[2].trim(); // Sender Email ID
    var subject = row[3].trim();
    
    //cc detals
    var cc_details = '';
    var cc_arr = [];

    if(row[15].trim() == '' && row[16].trim() == '' && row[17].trim() == '' && row[18].trim() == '' && row[19].trim() == '' 
       && row[20].trim() == '' && row[21].trim() == '' && row[22].trim() == '' && row[23].trim() == '' && row[24].trim() == '')
    {
      cc_details = ''
    }
    else
    {
      cc_arr[0] = row[15].trim(); //1
      cc_arr[1] = row[16].trim(); //2
      cc_arr[2] = row[17].trim(); //3
      cc_arr[3] = row[18].trim(); //4
      cc_arr[4] = row[19].trim(); //5
      cc_arr[5] = row[20].trim(); //6
      cc_arr[6] = row[21].trim(); //7
      cc_arr[7] = row[22].trim(); //8
      cc_arr[8] = row[23].trim(); //9
      cc_arr[9] = row[24].trim(); //10
      
      for (var i = 0; i < cc_arr.length; i++) 
      {
        if(cc_arr[i] != '')
        {
          if(cc_details == '')
          {
            cc_details = cc_arr[i];
          }
          else
          {
            cc_details = cc_details + ", " + cc_arr[i];
          }
        }
      }
    }
    
    //file attachment
    var file_details = "";
    var finalids = [];
    var attachments = [];
    var file = [];
 
    //var file = DriveApp.getFileById(id);
    
    if(row[25].trim() == '' && row[26].trim() == '' && row[27].trim() == '' && row[28].trim() == '' && row[29].trim() == '' )
    {
      file_details = ''
    }
    else
    {
      var filedata = '';
      var fileurl = [];
    
      var filename0 = (row[25].trim() != "") ? DriveApp.getFilesByName(row[25].trim()) : "";
      if(filename0 != "")
      {
          filedata = filename0.next();
          fileurl[0] = filedata.getUrl();
      }
      else
      {
        fileurl[0] = "";
      }
      
      var filename1 = (row[26].trim() != "") ? DriveApp.getFilesByName(row[26].trim()) : "";
      if(filename1 != "")
      {
          filedata = filename1.next();
          fileurl[1] = filedata.getUrl();
      }
      else
      {
        fileurl[1] = "";
      }
      
      var filename2 = (row[27].trim() != "") ? DriveApp.getFilesByName(row[27].trim()) : "";
      if(filename2 != "")
      {
          filedata = filename2.next();
          fileurl[2] = filedata.getUrl();
      }
      else
      {
        fileurl[2] = "";
      }
      
      var filename3 = (row[28].trim() != "") ? DriveApp.getFilesByName(row[28].trim()) : "";
      if(filename3 != "")
      {
          filedata = filename3.next();
          fileurl[3] = filedata.getUrl();
      }
      else
      {
        fileurl[3] = "";
      }
      
      var filename4 = (row[29].trim() != "") ? DriveApp.getFilesByName(row[29].trim()) : "";
      if(filename4 != "")
      {
          filedata = filename4.next();
          fileurl[4] = filedata.getUrl();
      }
      else
      {
        fileurl[4] = "";
      }
      
      finalids[0] = (fileurl[0] != "") ? DriveApp.getFileById(getIdFromUrl(fileurl[0])) : "";
      finalids[1] = (fileurl[1] != "") ? DriveApp.getFileById(getIdFromUrl(fileurl[1])) : "";
      finalids[2] = (fileurl[2] != "") ? DriveApp.getFileById(getIdFromUrl(fileurl[2])) : "";
      finalids[3] = (fileurl[3] != "") ? DriveApp.getFileById(getIdFromUrl(fileurl[3])) : "";
      finalids[4] = (fileurl[4] != "") ? DriveApp.getFileById(getIdFromUrl(fileurl[4])) : "";
          
      //   finalids[0] = (row[25].trim() != "") ? DriveApp.getFileById(getIdFromUrl(row[25].trim())) : "";
      //   finalids[1] = (row[26].trim() != "") ? DriveApp.getFileById(getIdFromUrl(row[26].trim())) : "";
      //   finalids[2] = (row[27].trim() != "") ? DriveApp.getFileById(getIdFromUrl(row[27].trim())) : "";
      //   finalids[3] = (row[28].trim() != "") ? DriveApp.getFileById(getIdFromUrl(row[28].trim())) : "";
      //   finalids[4] = (row[29].trim() != "") ? DriveApp.getFileById(getIdFromUrl(row[29].trim())) : "";
      
      //   file[0] = DriveApp.getFileById(finalids[0]);
      
      for (var i = 0; i < finalids.length; i++) 
      {
        if(finalids[i] != '')
        {
            // attachments.push(finalids[i].getAs(MimeType.PDF));
            attachments.push(finalids[i]);
        }
      }
    }
    
    const htmlTemplate = HtmlService.createTemplateFromFile("googlesheets.html");
    htmlTemplate.greeting = greeting;
    htmlTemplate.messagebody = messagebody + ".";
    htmlTemplate.para1 = para1 + (para1.length > 0  ? "." : "");
    htmlTemplate.para2 = para2 + (para2.length > 0  ? "." : "");
    htmlTemplate.para3 = para3 + (para3.length > 0  ? "." : "");
    htmlTemplate.para4 = para4 + (para4.length > 0  ? "." : "");
    
    const htmlforEmail = htmlTemplate.evaluate().getContent();
    console.log(htmlforEmail);
    
    // attachments: file.getAs(ContentType.PDF),
    if(cc_details != '')
    { 
      GmailApp.sendEmail(recipient, subject, "", {htmlBody: htmlforEmail,
                                                  cc: cc_details,
                                                  attachments: attachments,
                                                 });
    }
    else
    {
      GmailApp.sendEmail(recipient, subject, "send mails", {htmlBody: htmlforEmail,
                                                            attachments: attachments,
                                                           });
    }
    
  }
}


/*********************** CODE FOR SIDE BAR ***********************/
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
  var folderName = PropertiesService.getScriptProperties().getProperty('foldername');
  var folders = DriveApp.getFoldersByName(folderName);

  // var ui = SpreadsheetApp.getUi()
  // ui.alert(folderName);

   if(folderName == '')
    {
      PropertiesService.getScriptProperties().deleteAllProperties();
      Browser.msgBox('Please enter folder details.');
      
      return
    }

  var folder = folders.next();
  var file = folder.getFiles();

  // Logger.log('data : ',folname);

  Logger.log(file);
  // var file = DriveApp.getFiles();
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

function displayToast(name) {
  // var ui = SpreadsheetApp.getUi()
  // ui.alert(name);
  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().setProperty('foldername', name);
}

// function getLocation(val) {
//   var spreadsheet = SpreadsheetApp.getActive();
//   var sheetName = spreadsheet.getSheetName();
//   var selectedRange = spreadsheet
//     .getSelection()
//     .getActiveRange()
//     .getA1Notation();
      
//   spreadsheet.getRange(selectedRange).setValue(val).setFontColor('#00f'); 

// //  spreadsheet.getRange(selectedRange).setWrap(true);
  
//   spreadsheet.getRange(selectedRange).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

//   return {
//     sheet: sheetName,
//     range: selectedRange,
//     value: val
//   };
// }

function getLocation(val) {
  var spreadsheet = SpreadsheetApp.getActive();
//  var sheetName = spreadsheet.getSheetName();
  var selectedRange = spreadsheet
    .getSelection()
    .getActiveRange()
    .getA1Notation();
      
  spreadsheet.getRange(selectedRange).setValue(val).setFontColor('#00f').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); 

//  spreadsheet.getRange(selectedRange).setWrap(true);
  
//  spreadsheet.getRange(selectedRange).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  //  return {
  //    sheet: sheetName,
  //    range: selectedRange,
  //    value: val
  //  };
}





