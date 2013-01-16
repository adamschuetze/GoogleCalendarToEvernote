/*

A Google Apps script to extract Google Calendar entry notes, and send them to Evernote.  Install in your Google Drive, 
with Create>More>Script.  

You can change existing Google Calendar entry notes and the script will detect that, and send an updated entry to 
Evernote, which you can Merge with existing entries.

NOTES:

a) The first time the script runs, it configures itself.
b) Configure it to run on some sane time schedule (every 15 minutes, perhaps), with Resources>All your triggers.
c) It will only look at calendar entries between (now-daystoscan/2) to (now+daystoscan/2).  daystoscan is set to 14 
   days by default
d) It will only look at your default calendar, by default.  To make it look at a different calendar, edit the User 
   Properties (File>Project Properties>User properties)
e) Once it has configured itself and created the log spreadsheet, DO NOT delete that spreadsheet.  If you do, you need 
   to also delete the User Property "CalendartoEvernoteLogId".  Then running it again will make it reconfigure and 
   recreate a log sheet.

*/

function CalendartoEvernote () {
  var defaultNotebook = UserProperties.getProperty('defaultNotebook_calendar');
  var notebook='';
  if((defaultNotebook==undefined)||(defaultNotebook=='')) {
            notebook=' @' + 'Actions Pending';
  } else {
            notebook=' @' + defaultNotebook;
  }     
  var daystoscan=UserProperties.getProperty('daystoscan');
  if ((daystoscan==undefined)||(daystoscan=='')) {
    var daystoscan=14;
    UserProperties.setProperty('daystoscan',daystoscan);
  } 
  var CalendarId=UserProperties.getProperty('CalendarId');
  if ((CalendarId==undefined)||(CalendarId=='')) {
    var CalendarId=CalendarApp.getDefaultCalendar().getId();
    UserProperties.setProperty('CalendarId',CalendarId);
  }
  var CalendartoEvernoteLogId=UserProperties.getProperty('CalendartoEvernoteLogId');
  if ((CalendartoEvernoteLogId==undefined)||(CalendartoEvernoteLogId=='')) {
    createLogSheet();
    var CalendartoEvernoteLogId=UserProperties.getProperty('CalendartoEvernoteLogId');
  }
  var this_account=Session.getEffectiveUser().getEmail();
  var evernoteMail=UserProperties.getProperty('evernoteMail');
  var calendar=CalendarApp.getCalendarById(CalendarId);
  var now=new Date();
  var startdate=new Date(now.getTime()-daystoscan/2*86400000);
  var enddate=new Date(now.getTime()+daystoscan/2*86400000)
  var events=calendar.getEvents(startdate,enddate);
  var event_description='';
  var event_id='';
  var event_location='';
  var event_title='';
  var message_body='';
  var event_updated = new Number();
  var logsheet_timestamp = new Number();
  var num_events=events.length;
  for (var i=0;i<num_events;i++) {  
    event_description=events[i].getDescription();
    if (event_description!=''){
      event_id=events[i].getId();
      event_updated=Number(events[i].getLastUpdated());
      if (inLog(event_id,CalendartoEvernoteLogId)==true) {
        logsheet_timestamp=getLogTimeStamp(event_id,CalendartoEvernoteLogId);
        if (event_updated > logsheet_timestamp) { 
          event_location=events[i].getLocation();
          event_title=events[i].getTitle();
          logItem(event_id,event_title,event_updated,CalendartoEvernoteLogId,false);     
          message_body="Event ID: "+event_id
            +"<br>Sent to Evernote via: "+this_account
            +"<br>Title: "+event_title
            +"<br>Location: "+event_location
            +"<br>Notes: "+event_description;
          GmailApp.sendEmail(evernoteMail, event_title+notebook, '', {noReply:true, htmlBody: message_body});
        }
      } else {
        event_location=events[i].getLocation();
        event_title=events[i].getTitle();
        logItem(event_id,event_title,event_updated,CalendartoEvernoteLogId,true);
        message_body="Event ID: "+event_id
          +"<br>Sent to Evernote via: "+this_account
          +"<br>Title: "+event_title
          +"<br>Location: "+event_location
          +"<br>Notes: "+event_description;
        GmailApp.sendEmail(evernoteMail, event_title+notebook, '', {noReply:true, htmlBody: message_body});
      }
    }
  }
}

function getLogTimeStamp(event_id,CalendartoEvernoteLogId){
  var found = new Boolean(false);
  var sheet=SpreadsheetApp.openById(CalendartoEvernoteLogId).getSheets()[0];
  var num_events=sheet.getLastRow()-1;
  var IDs=sheet.getRange('A2:A'+num_events+1).getValues();
  var timestamps=sheet.getRange('C2:C'+num_events+1).getValues();
  var i=1;
  while ((i<=num_events)&&(found==false)){
    found=(event_id==IDs[i-1]);
    i++;
  }  
  return Number(timestamps[i-2]);
}

function inLog(event_id,CalendartoEvernoteLogId) {
  var found = new Boolean(false);
  var sheet=SpreadsheetApp.openById(CalendartoEvernoteLogId).getSheets()[0];
  var num_events=sheet.getLastRow()-1;
  if (num_events>=1) {
    var IDs=sheet.getRange('A2:A'+num_events+1).getValues();
    var i=1;
    while ((i<=num_events)&&(found==false)){
      found=(event_id==IDs[i-1]);
      i++;
    }   
  }
  return found;
}

function logItem(event_id,event_title,event_updated,CalendartoEvernoteLogId,new_item) {
  var sheet=SpreadsheetApp.openById(CalendartoEvernoteLogId).getSheets()[0];
  if (new_item==false){
    var num_events=sheet.getLastRow()-1;
    if (num_events>=1){
      var IDs=sheet.getRange('A2:A'+num_events+1).getValues();
      var i=1;
      var found=new Boolean(false);
      while((i<=num_events)&&(found==false)){
        found=(event_id==IDs[i-1]);
        i++;
      }
    }
    sheet.deleteRow(i);
  }
  sheet.appendRow([event_id,event_title,event_updated]);
}

function createLogSheet()
{
  var spreadsheet = SpreadsheetApp.create("Google Calendar to Evernote log");
  var sheet = spreadsheet.getSheets()[0];
  UserProperties.setProperty('CalendartoEvernoteLogId', spreadsheet.getId());
  sheet.appendRow(['CalendarItemId','CalendarItemTitle','CalendarItemTimestamp']);
}
