var folderId = "/*folder id*/";
var fileName = "test1";
var newRecordingSheet;
var sheets = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Failed"];
/**
 * This function is called before parsing any of the classes from the original sheet
 * 
 * Create the new sheet file, add headers, and add sheets for each day of the week.
 */
function initializeSheet(bulkRecordingSheetArg) {
  //Create sheet, headers, and sheets for each day of the week
  var newSheetId = createSpreadSheetInFolder();//call to create spreadsheetinfolder
  newRecordingSheet = SpreadsheetApp.openById(newSheetId); 
  newRecordingSheet.insertSheet("Monday");
  newRecordingSheet.insertSheet("Tuesday");
  newRecordingSheet.insertSheet("Wednesday");
  newRecordingSheet.insertSheet("Thursday");
  newRecordingSheet.insertSheet("Friday");
  newRecordingSheet.insertSheet("Failed");
  var sheet = newRecordingSheet.getSheetByName('Sheet1');
  newRecordingSheet.deleteSheet(sheet);
  
  for(var i = 0; i < 5; i++){
    sheet = newRecordingSheet.getSheetByName(sheets[i]);
    sheet.getRange("A3:P3").setBackgroundRGB(211,211,211);
    sheet.getRange("A3:P3").setFontWeight('bold');
    sheet.getRange("A1").setValue("Date");
    sheet.getRange("A3").setValue("Course");
    sheet.getRange("B3").setValue("Record Start 24h")
    sheet.getRange("C3").setValue("Instructor");
    sheet.getRange("D3").setValue("Room");
    sheet.getRange("E3").setValue("Start Time");
    sheet.getRange("F3").setValue("Recording Started");
    sheet.getRange("G3").setValue("Uploaded to Kaltura");
    sheet.getRange("H3").setValue("Added to Canvas");
    sheet.getRange("I3").setValue("Canvas Link");
    sheet.getRange("J3").setValue("Not In Class");
    sheet.getRange("K3").setValue("Embed Code");
    sheet.getRange("L3").setValue("Video Location");
    sheet.getRange("M3").setValue("Module Number or Assignment Number");
    sheet.getRange("N3").setValue("Prefix");
    sheet.getRange("O3").setValue("Suffix");
    sheet.getRange("P3").setValue("Publish");
  }
  bulkRecordingSheetArg.getRange("J3").setValue(newRecordingSheet.getUrl());
  bulkRecordingSheetArg.getRange("J2").uncheck();
}

/**
 * Called from initializeSheet Function
 * Creates new sheet with filename declared at the top of this page
 * Stores it in the folder declared at the top of this page
 * 
 * You can change the file name and folder id at the top to change the name of the file after creation
 * and to change the folder that it's stored in
 */
function createSpreadSheetInFolder() {
    var ss_new = SpreadsheetApp.create(fileName);
    var ss_new_id = ss_new.getId();
    var newfile = DriveApp.getFileById(ss_new_id);
    newfile.moveTo(DriveApp.getFolderById(folderId))
    return ss_new_id;
}

/**
 * This function will be called within each successful loop of the Bulk Recording Scheduler
 * 
 * Fill in a row of the spreadsheet with the class information
 */
function createSchedule(eventObject) {
  //use the information parsed from the sheet to fill the cells with the correct information per class.
  //let tmp = DriveApp.getFileById(bulkRecordingSheet.getRange("J3").getValue());
  newRecordingSheet = SpreadsheetApp.openByUrl(bulkRecordingSheet.getRange("J3").getValue());
  var daysComp = ["m", "t", "w", "r", "f"];
  if(!eventObject.evRecurrence){// Missing days info, this is required info
  }
  else{
    for(var i = 0; i < 5; i++){//for each day/tab in the sheet
      if(eventObject.evRecurrence.includes(daysComp[i])){//check if this class occurrs on that day
        sheet = newRecordingSheet.getSheetByName(sheets[i]);
        //Here's where we will fill in the information for the class.
        var link = getRoomLink(eventObject.roomNumber);
        if(link != -1){//if the room number matched one that we can record
          sheet.appendRow([eventObject.eventTitle, eventObject.eventStartTime, eventObject.eventInstruct, , eventObject.eventTime]);
          var richValue = SpreadsheetApp.newRichTextValue().setText(eventObject.roomNumber).setLinkUrl(link).build();
          sheet.getRange(sheet.getLastRow(), 4).setRichTextValue(richValue);
          sheet.getRange(sheet.getLastRow(), 6).insertCheckboxes();
          sheet.getRange(sheet.getLastRow(), 7).insertCheckboxes();
          sheet.getRange(sheet.getLastRow(), 8).insertCheckboxes();
        }else{ // Room number did not match one that we can record
          sheet = newRecordingSheet.getSheetByName("Failed");
          sheet.appendRow([eventObject.eventTitle, , eventObject.eventInstruct, eventObject.roomNumber, eventObject.eventTime]);
          sheet.getRange(sheet.getLastRow(),1,1,8 ).setBackgroundRGB(224, 102, 102);
        }
      }
    }
  }
}

/**
 * This function will be called after the loop is done. 
 * 
 * Reorder the classes that are in here by the time that they start. 
 * there is a column that will be start time in 24 hr format, we will then sort from that. 
 */
function organizeSchedule(){
  //reorder the entries in the sheet by time.
  newRecordingSheet = SpreadsheetApp.openByUrl(bulkRecordingSheet.getRange("J3").getValue());
  for(var i = 0; i < 6; i++){
    sheet = newRecordingSheet.getSheetByName(sheets[i]);
    range = sheet.getRange("A4:P");
    range.sort(2);
  }
}

/**
 * This function is called while filling in rows in order to get the link to the live stream
 * 
 * It is just a switch that returns either the link or a -1 if the link DNE
 */
function getRoomLink(roomNum){
  switch(roomNum){
    case 2211:
          return "/*link to livestream*/";
    case 2260:
          return "/*link to livestream*/";
    case 3250:
          return "/*link to livestream*/";
    case 5229:
          return "/*link to livestream*/";
    case 5240:
          return "/*link to livestream*/";
    case 5246:
          return "/*link to livestream*/";
    case "Lubar":
          return "/*link to livestream*/";
    case 7200:
          return "/*link to livestream*/";
    case 2225:
          return "/*link to livestream*/";
    case 3226:
          return "/*link to livestream*/";
    case 3247:
          return "/*link to livestream*/";
    case 3253:
          return "/*link to livestream*/";
    case 3260:
          return "/*link to livestream*/";
    case 3261:
          return "/*link to livestream*/";
    case 3268:
          return "/*link to livestream*/";
    case 5223:
          return "/*link to livestream*/";
    default:
            return -1;
  }
}