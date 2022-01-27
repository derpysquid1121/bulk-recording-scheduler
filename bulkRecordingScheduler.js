//Load Moment Library
var moment = Moment.load();
var calChoice = "/*main calendar link*/";
var calChoice2 = "/*A second calendar if needed*/";
var bulkRecordingSheet;

var GLOBAL = {
  calendarVar :"This will be set in the switch",
  formMap : {
    eventTitle : "Event Title",
    startTime: "Event Start Time",
    endTime: "Event End Time",
    evRecurrence: "Days to repeat on",
    roomNumber: "Room Number",
    description: "description",
    ipAddress: "IP Address",
    calendarChoice: "Which calendar to use",
  },
}

function parseData() { //Parse the data of each row and call the function to create a calendar event. 
  var sheet = SpreadsheetApp.getActive().getSheetByName("Classes"); //get the active spreadsheet called classes
  bulkRecordingSheet = sheet;
  var data = sheet.getDataRange().getValues(); //get the data from the sheet

  sheet.getRange("Classes!N:P").uncheck(); //uncheck all the error logging boxes so they can be rechecked if there are still errors 
  sheet.getRange("Classes!B2:N").setBackgroundRGB(255,255,255); //recolor cells as white so that the error area can be colored again


  var days;      //array that holds the days of the week that this class is held on.
  var title;     //name of the class
  var roomNum;   //the number of the room to record
  var time;      //the time that the class starts
  var startDate; //date that classes start on
  var endDate;   //date that classes end on
  var missingData; //needed data is missing = 1, all needed data included = 0
  var instruct; //the name of instructor. Needed for creating the schedule sheet

  if(sheet.getRange("J2").getValue()){ // if true that means we want a new sheet
    sheet.getRange("J3").setValue("");
    initializeSheet(sheet);
  }
  else{ // This means we use the previously created one
    try{
      var url = sheet.getRange("J3").getValue(); //get the url
      if(url){ //if there is a value in url
        SpreadsheetApp.openByUrl(url); //attempt to open the spreadsheet by url
      }else{throw Exception;} //if no value is in url, then throw an exception to be caught by the catch block
    }catch(Exception){ //if the previous url doesn't exist or doesn't work, we end up here
      initializeSheet(sheet);
    }
  }
  
  data.forEach(function (col, index) {
    var spreadSheetInfo = { //Object to hold the data needed for the second spreadsheet
      eventTitle : "Event Title",
      eventInstruct : "Event Instructor",
      eventTime: "Event Time",
      eventStartTime: "event start time 24h",
      evRecurrence: "Days to repeat on",
      roomNumber: "Room Number",
      ipaddr: "IP Addr",
    };
    if(index == 0){ //There is nothing in this row
      //nothing
    }
    else if(sheet.getRange(index + 1, 17).getValue() == false){//if the row has not been added to the calendar. 
      missingData = 0; // set the missing data state variable
      if(index == 1){ // location of classes start date
        startDate = col[1];
        if(!startDate){ // if start date is missing
          missingData = 1;
          sheet.getRange(index + 1, 2).setBackgroundRGB(224, 102, 102);
        }
      }
      if(index == 2){ // location of classes end date 
        endDate = col[1];
        if(!endDate){ // if end date is missing 
          missingData = 1;
          sheet.getRange(index + 1, 2).setBackgroundRGB(224, 102, 102);
        }
      }
      if(index>2){ //Title and Empty Rows
      //-----------------title------------------
        title = col[1];
        spreadSheetInfo.eventTitle = title;
        if(!title){ //if title is missing
          missingData = 1;
          sheet.getRange(index + 1, 2).setBackgroundRGB(224, 102, 102);
          spreadSheetInfo.eventTitle = "missing title";
        }
        //----------------------------------------
        //--------------Instructor----------------
        instruct = col[3];
        spreadSheetInfo.eventInstruct = instruct;
        if(!instruct){ //if instructor name is missing
          instruct = "missing instructor name";
        }
        //----------------------------------------
        //--------------days----------------------
        var daysOriginal = col[4];
        if(!daysOriginal){ //if days is missing
          missingData = 1;
          sheet.getRange(index + 1, 5).setBackgroundRGB(224, 102, 102);
        }
        else{
          days = daysOriginal.split(""); //split days into their separate char's
          spreadSheetInfo.evRecurrence = daysOriginal.toLowerCase().split("");
          }
        //----------------------------------------
        //--------------roomnum-------------------
        roomNum = col[6];
        spreadSheetInfo.roomNumber = roomNum;
        if(!roomNum){ //if roomnum is missing
            missingData = 1;
            sheet.getRange(index + 1, 7).setBackgroundRGB(224, 102, 102);
            spreadSheetInfo.roomNumber = "Missing room Number";
        }

        //----------------------------------------
        //_.~"~._.~"~._.~"~._.~"~.__.~"~._.~"~._.~"~._.~"~._VV-PARSE TIME-VV_.~"~._.~"~._.~"~._.~"~.__.~"~._.~"~._.~"~._.~"~._
        time = col[5];
        spreadSheetInfo.eventTime = time;
        if(!time){ //if time is missing
          missingData = 1;
          sheet.getRange(index + 1, 6).setBackgroundRGB(224, 102, 102);
          spreadSheetInfo.eventTime = "Missing time";
        }
        else if(time.length > 9 || time.length < 7){ //if time is too long or too short to be in the normal time format ie. 1000-1100 or 100-200.
          missingData = 1;
          sheet.getRange(index + 1, 6).setBackgroundRGB(224, 102, 102);
          spreadSheetInfo.eventTime = "Missing time"; 
        }
        else{//TODO: parse time here
          var timeOffsetBefore=2;//offset before event time
          var timeOffsetAfter=5;//offset after event time
          var dayPeriodStart;  //am or pm for start time
          var dayPeriodEnd;   //am or pm for end time 
          var startHour;     //hour that recording starts
          var startMin;     //min that recording starts
          var endHour;     //hour that recording ends
          var endMin;     //min that recording ends
          var times;     //array that holds the beginning and end times in original format 
          
          times = time.split('‑'); //break up start and end time in original format
          if(times.length != 2){
            times = time.split('-');
          }

          //-------------Get hour and min of start time-----------------
          if(times[0].length == 3){ //if start time is 3 chars
            startHour = times[0].charAt(0);
            startMin = times[0].substring(1,3);
          }else if(times[0].length == 4){ //if start time is 4 chars
            startHour = times[0].substring(0, 2);
            startMin = times[0].substring(2,4);
          }
          if(startMin.charAt(0) == 0){ //strip leading zero from minute for use in calculating offset
            startMin = startMin.substring(1);
          }
          //------------------------------------------------------------
          //---------------Get hour and min of end time-----------------
          if(times[1].length == 3){ //if end time is 3 chars
            endHour = times[1].charAt(0);
            endMin = times[1].substring(1,3);
          }else if(times[1].length == 4){ //if end time is 4 chars
            endHour = times[1].substring(0, 2);
            endMin = times[1].substring(2,4);
          }
          if(endMin.charAt(0) == 0){ //strip leading zero from minute for use in calculating offset
            endMin = endMin.substring(1);
          }
          //------------------------------------------------------------
          //------------Offset start time to begin before class---------
          //decrement minutes by offset -> if(minutes is negative){then decrement hour and set minutes = 60 - abs(minutes)}
          var temp = startMin;
          temp = temp - timeOffsetBefore;
          if(temp < 0){ //if we need to drop back an hour for the offset
            if(startHour == 1){ //if dropping back an hour makes us switch from 1 to 12
              startHour = 12;
            }
            else{//otherwise we can just decrement hour by 1
              startHour--;
            }
            startMin = 60 - Math.abs(temp);//find the minute to start the class in the new hour
          }
          else{
            startMin = temp;
          }
          //-----------------------------------------------------------
          //------------Offset end time to end after class-------------
          //decrement minutes by offset -> if(minutes is negative){then decrement hour and set minutes = 60 - abs(minutes)}
          var temp = endMin;
          temp = +temp + +timeOffsetAfter;
          if(temp >= 60){ //if we need to go forward an hour for the offset
            if(endHour == 12){ //if dropping back an hour makes us switch from 1 to 12
              endHour = 1;
            }
            else{//otherwise we can just decrement hour by 1
              endHour++;
            }
            endMin = Math.abs(temp) - 60;//find the minute to start the class in the new hour
          }
          else{
            endMin = temp;
          }
          //Logger.log(startHour + ":" + startMin + "-" + endHour + ":" + endMin);
          //-----------------------------------------------------------
          //------------Figure out am vs pm for start time-------------
          if(startHour >= 7 && startHour <= 11){ //between 8 and 11 inclusive is am
            dayPeriodStart = "am";
            spreadSheetInfo.eventStartTime = startHour.toString().concat(startMin);
          }
          else if(startHour == 12){ //the hour of 12 is always pm
            dayPeriodStart = "pm";
            spreadSheetInfo.eventStartTime = startHour.toString().concat(startMin);
          }
          else if(startHour >= 1 && startHour <= 6){ //between 1 and 6 inclusive is pm
            dayPeriodStart = "pm";
            spreadSheetInfo.eventStartTime = startHour + 12;
            spreadSheetInfo.eventStartTime = spreadSheetInfo.eventStartTime.toString().concat(startMin);
          }
          //-----------------------------------------------------------
          //------------Figure out am vs pm for end time---------------//TODO: May need to make changes to this if classes ever end after 8 pm
          if(endHour >= 8 && endHour <= 11){ //between 8 and 11 inclusive is am
            dayPeriodEnd = "am";
          }
          else if(endHour == 12){ //the hour of 12 is always pm
            dayPeriodEnd = "pm";
          }
          else if(endHour >= 1 && endHour <= 7){ //between 1 and 7 inclusive is pm
            dayPeriodEnd = "pm";
          }
          //-----------------------------------------------------------
          if(startMin < 10){ //return leading zero removed in the blocks where we get the time from the raw data
            startMin = "0" + startMin;
          }
          if(endMin < 10){ //return leading zero removed in the blocks where we get the time from the raw data
            endMin = "0" + endMin;
          }
        }
        Logger.log(startHour+dayPeriodStart + "-" + endHour+dayPeriodEnd);

        //create Date String with information passed in. 
        var year = startDate.getFullYear();
        var month = startDate.getMonth()+1;
        var day = startDate.getDate();
        var timeZoneOffset = new Date().getTimezoneOffset()/60;

        var recordingStartDate = moment(month + " " + day + " " + year + ", " + startHour + ":" + startMin + ":00 " + dayPeriodStart);
        var recordingEndDate = moment(month + " " + day + " " + year + ", " + endHour + ":" + endMin + ":00 " + dayPeriodEnd);
        
        //_.~"~._.~"~._.~"~._.~"~.__.~"~._.~"~._.~"~._.~"~.__.~"End of Parsing Time"~.__.~"~._.~"~._.~"~._.~"~.__.~"~._.~"~._.~"~._.~"~._

        //-----------vvvv---EITHER PASS DATA INTO FUNCTION TO CREATE EVENT OR INDICATE THAT ROW DATA INSUFFICIENT--vvvvvvvvvvvvv--------------vvvvvvvvvvvvvvvvv----
        if(!missingData){ //if we have all the data that we need.
          var recurrence = getRecurrence(days, endDate);//generate the recurrence used to make the calendar event
          var eventObject = getInfo(roomNum.toString(10), title, recordingStartDate, recordingEndDate); //generate the object with all the info needed to make the calendar event
          if(eventObject == -1 || recurrence == -1){//if making the object or reccurrence failed. 
              Logger.log("Parsing error: Roomnum error?: " + eventObject + " Recurrence Error?: " + recurrence + " : -1 is where the error originated");
              if(eventObject == -1){//if error occurred during parsing of options with room number, check room number 
                sheet.getRange(index + 1, 15).insertCheckboxes();
                sheet.getRange(index + 1, 15).check();
                } 
              if(recurrence == -1){//if error occurred during parsing of days(recurrence), check recurrence 
                sheet.getRange(index + 1, 16).insertCheckboxes();
                sheet.getRange(index + 1, 16).check();
                } 
          }
          else{//successful object created
            var eventSeries = getEventSeries(eventObject, recurrence); //create the calendar event
            if(eventSeries != -1){ //if creating the calendar event was successful
              Logger.log(spreadSheetInfo.evRecurrence + " " + spreadSheetInfo.eventInstruct + " " + spreadSheetInfo.eventTime + " " + spreadSheetInfo.eventTitle + " " + spreadSheetInfo.roomNumber);
              sheet.getRange(index + 1, 14, 1, 4).insertCheckboxes();
              sheet.getRange(index + 1, 14).uncheck();
              sheet.getRange(index + 1, 15).uncheck();
              sheet.getRange(index + 1, 16).uncheck();
              sheet.getRange(index + 1, 17).check();
              createSchedule(spreadSheetInfo); // function call to fill info on schedule sheet. 
            }
            else{//creating the calendar event failed, highlight entire row. 
              sheet.getRange(index + 1,2,1,12 ).setBackgroundRGB(224, 102, 102);
              Logger.log("eventSeries Creation Failed");
            }
          }
        }else { //if the row had incorrect or missing data
          sheet.getRange(index + 1, 14).insertCheckboxes();
          sheet.getRange(index + 1, 14).check();
        }
        //--------^^^^^^^^^^^----------------^^^^^^^^^^^^------------------^^^^^^^^^^^^^^^-----------------^^^^^^^^^^^^^^^^--------------------
      }
    }
    else{ //Row was marked as successfully being created
      Logger.log("Row " + index + " added Already");
    }
  });
  organizeSchedule();
}

function getRecurrence(days, endDate){
  var newEndDate = new Date(endDate); //put end date into date format
  var newDays = []; //create array so that we can make onlyOnWeekday rule with days data

  for(var i = 0; i<days.length; i++){ //for each day in the days array for a class add calendarapp weekday to new array
    switch (days[i])  {
        case "M":
        case "m":
            newDays.push(CalendarApp.Weekday.MONDAY);
            break;
        case "T":
        case "t":
            newDays.push(CalendarApp.Weekday.TUESDAY);
            break;
        case "W":
        case "w":
            newDays.push(CalendarApp.Weekday.WEDNESDAY);
            break;
        case "R":
        case "r":
            newDays.push(CalendarApp.Weekday.THURSDAY);
            break;
        case "F":
        case "f":
            newDays.push(CalendarApp.Weekday.FRIDAY);
            break;
        default:
            return -1;
    }
  }
  //create and then return the recurrence to be used when making the event. 
  var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(newDays).until(newEndDate);
  return recurrence;
}

function getInfo(roomNum, title, startTime, endTime){
  var eventObject = {};
  eventObject.eventTitle = title;
  eventObject.startTime = startTime;
  eventObject.endTime = endTime;
  eventObject.calendarChoice = calChoice;
  switch(roomNum){
    case "2211":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "2260":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle =  eventObject.eventTitle;
          eventObject.calendarChoice = calChoice2;
          break;
    case "3250":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.eventTitle = eventObject.eventTitle;
          eventObject.calendarChoice = calChoice2 
          break;
    case "5229":
          eventObject.ipAddress = "/*info for event*/"; 
          eventObject.eventTitle = eventObject.eventTitle;
          eventObject.calendarChoice = calChoice2
          break;
    case "5240":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "5246":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
    case "Lubar":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "7200":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "2225":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "3226":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "3247":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;     
    case "3253":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "3260":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "3261":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "3268":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/";
          break;
    case "5223":
          eventObject.ipAddress = "/*info for event*/";
          eventObject.description = "/*info for event*/";
          eventObject.eventTitle = "/*info for event*/r";
          break;
    default:
            return -1;
  }
  return eventObject;
}

function getEventSeries(eventObject, recurrence){
  Logger.log(eventObject.calendarChoice);
  try{
    var options = {
      description : eventObject.description,//This depends on room number 
      location : eventObject.ipAddress, //This depends on room number
      }
    var eventSeries = CalendarApp.getCalendarById(eventObject.calendarChoice).createEventSeries(
      eventObject.eventTitle,
      new Date(eventObject.startTime),
      new Date(eventObject.endTime),
      recurrence,
      options
      );
    return eventSeries;
  }catch(Exception){
    Logger.log(Exception);
    return -1;
  }
  //this creates an event series
}
