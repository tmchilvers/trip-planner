//  gsheets-trip: Google Sheets Apps Script custom functions for working
//  with dynamic trip-planning functionalities
//
//  Designed for personal use to help you plan all of your fun trips!
//
//  Author: Tristan Chilvers (https://www.tristanchilvers.com/)
//
//  The easiest way to use yourself is to use File > Make a copy on this Sheet:
//  https://docs.google.com/spreadsheets/d/1B_dnR0HjFd-yClSEWd2_BwNsqBRuEJXZhJbATznj11U/edit?usp=sharing
//
//  Code on GitHub: https://github.com/tmchilvers/trip-planner.git
//  (If you made a copy, check above link for the most updated version).
//
//
//  This is open source software that is free to use and share, as covered by the
//  MIT License.
//  
//  V.0.1.1

//  =====================================================================================
//                                    CONSTANTS
//  =====================================================================================

//  -------------------------------------------------------------------------------------
//  Sheets

var SHEET_TRIP_DETAILS = SpreadsheetApp.getActive().getSheetByName('*Trip Details');
var SHEET_TRAVELERS = SpreadsheetApp.getActive().getSheetByName('*Travelers');
var SHEET_EVENTS = SpreadsheetApp.getActive().getSheetByName('*Events');
var SHEET_CALENDAR = SpreadsheetApp.getActive().getSheetByName('Calendar');
var SHEET_FINANCE = SpreadsheetApp.getActive().getSheetByName('Finance');
var SHEET_STATIC_DATA = SpreadsheetApp.getActive().getSheetByName('StaticData');

//  -------------------------------------------------------------------------------------
//  Ranges

//  Trip Details Cells
var TRIP_DAYS = SHEET_TRIP_DETAILS.getRange(5, 3, 1, 1).getValue();
var NUM_EVENTS = SHEET_TRIP_DETAILS.getRange(8, 3, 1, 1).getValue();

//  Events Cells
var RANGE_EVENTS_COLUMNS = SHEET_EVENTS.getRange(4, 1, NUM_EVENTS + 4, SHEET_EVENTS.getMaxColumns());

//  Calendar Cells
var RANGE_CALENDAR_TITLE = SHEET_CALENDAR.getRange(1, 1, 1, SHEET_CALENDAR.getMaxColumns());
var RANGE_CALENDAR_DATE = SHEET_CALENDAR.getRange(4, 2, 1, SHEET_CALENDAR.getMaxColumns());
var RANGE_CALENDAR_TIME = SHEET_CALENDAR.getRange(6, 1, 29, 1);

//  Static Data Cells
var RANGE_STATIC_DATA_EVENTS_RANGE = SHEET_STATIC_DATA.getRange(1, 1, SHEET_STATIC_DATA.getMaxRows(), 1);
var RANGE_STATIC_DATA_TIME_START = SHEET_STATIC_DATA.getRange(4, 6, SHEET_STATIC_DATA.getMaxRows(), 1);
var RANGE_STATIC_DATA_TIME_END = SHEET_STATIC_DATA.getRange(4, 7, SHEET_STATIC_DATA.getMaxRows(), 1);

//  -------------------------------------------------------------------------------------
//  UI

//  Event Colors
var colors = ["#e06666", "#6fa8dc","#fe91ff","#75e192","#fffd94","#ffb470"]

//  Trip Details
var distanceUnit = SHEET_TRIP_DETAILS.getRange(10,3,1,1).getValues();

//  Event Columns Data
var EVENT_NAME = 0;
var BOUGHT_FROM = 1;
var EVENT_LINK = 2;
var EVENT_COST = 3;
var OVERNIGHT = 4;
var SHOW_EVENT = 5;
var DATE_START = 6;
var DATE_END = 7;
var TIME_START = 8;
var TIME_END = 9;
var SPLIT_COST = 10;
var EVENT_CHARGE_TRAVELERS = 11;
var EVENT_TYPE = 12;
var EVENT_DISTANCE = 13;
var EVENT_NOTES = 14;
var EVENT_OVERLAP = 15;
var EVENT_INCORRECT_TIME = 16;

//  =====================================================================================
//                                   GLOBAL VARIABLES
//  =====================================================================================

//  List of Cells for each event (this allows the script to merge and unmerge event cells)
var eventRanges = [];

//  =====================================================================================
//                                      FUNCTIONS
//  =====================================================================================

//  -------------------------------------------------------------------------------------
//                                   HELPER FUNCTIONS
//  -------------------------------------------------------------------------------------

//  -------------------------------------------------------------------------------------
//  Grab the color associated with the event (argument takes string of the event type)

function getColor(event) {
  if(event == "Transportation") return colors[0];
  else if(event == "Lodging") return colors[1];
  else if(event == "Food") return colors[2];
  else if(event == "Activity") return colors[3];
  else if(event == "Shopping") return colors[4];
  else if(event == "MISC") return colors[5];
}

//  -------------------------------------------------------------------------------------
//  The global "Update" function - if the dates of the trip is changed, change the calendar layout
//  Takes a cell as an argument

function onEdit(e) {
    if (e.range.columnStart == 3 && e.range.rowStart == 12 && sh.getName() == "*Trip Details") {
      distanceUnit = e.valueOf();
    }  

    var sh=e.range.getSheet();

    if (e.range.columnStart == 3 && e.range.rowStart == 4 && sh.getName() == "*Trip Details" || 
        e.range.columnStart == 3 && e.range.rowStart == 5 && sh.getName() == "*Trip Details" ) {
      clearCalendar();
      showCols();
      hideCols();
      mergeCalendarTitle();
      displayEvents();
    }
}

//  -------------------------------------------------------------------------------------
//                                 CALENDAR CORE FUNCTIONS
//  -------------------------------------------------------------------------------------

//  -------------------------------------------------------------------------------------
//  Updates the calendar events
//  - Clears the calendar data
//  - Unmerge cells (by storing which cells were merged) ** IF IT LOSES THAT DATA IT HAS TO BE MANUALLY UNMERGED **
//  - Add each event data to the calendar
//  - Merge cells for each event (and store which cells were merged)
//  
//  * Called by button by user *

function updateCalendar() {
  clearCalendar();
  displayEvents();
}

//  -------------------------------------------------------------------------------------
//  Resets the visuals/data of calendar events
//  - Removes all event data
//  - Breaks apart all merged cells
//  - Resets static data containing the list of event ranges

function clearCalendar() {
    //  Reset the list of event ranges
    eventRanges = [];

    //  Grab the data from static sheet that contains the list of cells that each event occupies
    var eventsCells = RANGE_STATIC_DATA_EVENTS_RANGE.getValues();

    //  Store the range of cells for each event
    var row = 0;
    while(eventsCells[row][0] != '') {
        
        //  Save the range
        eventRanges.push(eventsCells[row][0]);
        
        //  Remove the data for that event from the static sheet
        SHEET_STATIC_DATA.getRange(row + 1, 1, 1, 1).setValue('');
        row++;
    }
    
    //  Break apart all merged cells of the calendar (if there were any in the first place)
    if(row != 0) {
      eventRanges.forEach(function(value) { 
        SHEET_CALENDAR.getRange(value).breakApart();
      });
    }

    //  Clear cell colors and reset the alternating color background of the entire calendar
    for(var i = 6; i < 30; i++) {
        for(var j = 2; j < TRIP_DAYS + 2; j++) {

          SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(i,j,1,1).getA1Notation()).clearContent();
          
          if(i%2)
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(i,j,1,1).getA1Notation()).setBackground('#f3f3f3');
          else
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(i,j,1,1).getA1Notation()).setBackground('white');                        
        }
    }
}

//  -------------------------------------------------------------------------------------
//  Display all events on the calendar sheet
//  - Grab data from each event
//  - Set visuals for all cells of the duration of the event
//  - Store the range of cells for each event to allow for easy merge/unmerge of event cells

function displayEvents() {
    //  Reset the list of event ranges
    eventRanges = [];

    //  Grab the data from static sheet that contains the time start/end rounded
    var timeStartsRounded = RANGE_STATIC_DATA_TIME_START.getValues();
    var timeEndsRounded = RANGE_STATIC_DATA_TIME_END.getValues();    

    //  Grab values of events and calendar dates/times
    var eventsValues = RANGE_EVENTS_COLUMNS.getValues(); // at first event
    var calendarDateValues = RANGE_CALENDAR_DATE.getValues();
    var calendarTimeValues = RANGE_CALENDAR_TIME.getValues();
    
    //  Instantiate event variables to grab data for each event
    var eventName, overnight, timeStart, timeEnd;
    var dateStart = new Date();
    var dateEnd = new Date();

    //  -------------------------------------------------------------------------------------
    //  Go through all events
    for (var i = 0; i < NUM_EVENTS; i++) { 
      
        //  If the event is to be displayed on the calendar
        if(eventsValues[i][SHOW_EVENT] == 'Yes') {

          //  Grab all data associated with the event
          eventName = eventsValues[i][EVENT_NAME];
          overnight = eventsValues[i][OVERNIGHT];
          dateStart = eventsValues[i][DATE_START];
          dateEnd = eventsValues[i][DATE_END];
          timeStart = timeStartsRounded[i][0];
          timeEnd = timeEndsRounded[i][0];
          eventType = eventsValues[i][EVENT_TYPE];
          eventDistance = eventsValues[i][EVENT_DISTANCE];
          eventOverlap = eventsValues[i][EVENT_OVERLAP];
          eventIncorrectTime = eventsValues[i][EVENT_INCORRECT_TIME];

          //  Don't display event if it collides with another event or has incorrect time
          if (eventOverlap || eventIncorrectTime) { break }

          //  Set index variables for looping through data
          var col = row = 0;
          
          //  Calculate number of days for the event
          var days = (dateEnd.valueOf() - dateStart.valueOf())/(24*60*60*1000);
          
          //  -------------------------------------------------------------------------------------
          // Go through all Dates on the calendar
          for(var j = 0; j < calendarDateValues[0].length - 1; j++) {
          
            // Find column for Date Start on the calendar
            if(calendarDateValues[0][j].valueOf() == dateStart.valueOf()) {
              col = j + 2; // the date
              break;
            }
          }

          //  -------------------------------------------------------------------------------------
          // Go through all Times on the calendar
          for(var j = 0; j < 24; j++) {
            
            //  Find row for Time Start on the calendar
            if(calendarTimeValues[j][0].valueOf() == timeStart.valueOf()) {
              row = j + 6; // the time
              break;
            } 
          }

          //  Set index variables for looping through data
          var rowStart = row;

          //  Set visuals for the event on the calendar
          if(eventType == "Transportation" && eventDistance > 0) {
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(row,col,1,1).getA1Notation()
                                    ).setBackground(getColor(eventType)
                                    ).setValue(eventName + "\n" + eventDistance + " " + distanceUnit
                                    ).setFontWeight("bold"
                                    ).setFontStyle("normal");
          }
          
          else {
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(row,col,1,1).getA1Notation()
                                    ).setBackground(getColor(eventType)
                                    ).setValue(eventName
                                    ).setFontWeight("bold"
                                    ).setFontStyle("normal");
          }
          //  -------------------------------------------------------------------------------------
          //  Go through all Times on the calendar and set visuals for the duration of the event
          for(var j = row - 5; j < 24; j++) {
            row = j + 6
            
            //  Find row for Time End on the calendar and exit
            if(calendarTimeValues[j][0].valueOf() == timeEnd.valueOf()) {
              row -= 1;
              break;        
            }

            //  Set color
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(row,col,1,1).getA1Notation()).setBackground(getColor(eventType)); 
          }

          //  Store the cell range for the event as static data (this allows the script to merge/unmerge the cells on the calendar)
          currEvent = SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(rowStart,col,(row + 1) - rowStart,1).getA1Notation());
          eventRanges.push(currEvent);

          //  -------------------------------------------------------------------------------------
          //  For the overnight events, it needs to span multiple dates on the calendar
          if(overnight == "Yes") {

            // Fill all days until the final day of the event
            for(var j = 1; j < days; j++) {
              
              //  Set title of event as a continuation of the event title
              SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(6,col + j,1,1).getA1Notation()
                                      ).setValue("(" + eventName + ")"
                                      ).setFontWeight("normal"
                                      ).setFontStyle("italic");   

              //  Set color for the duration of the event
              for(var k = 0; k < 24; k++) {      
                row = k + 6;      
                SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(row,col + j,1,1).getA1Notation()).setBackground(getColor(eventType));             
              }

              //  Store the cell range of the event as static data
              currEvent = SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(6, col + j, 24, 1).getA1Notation());
              eventRanges.push(currEvent);
            }

            //  Set column to the last day of the event
            col += days;

            //  Set title of event as continuation of title
            SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(6,col,1,1).getA1Notation()
                                    ).setValue("(" + eventName + ")"
                                    ).setFontWeight("normal"
                                    ).setFontStyle("italic");   

            //  -------------------------------------------------------------------------------------
            //  Go through all times on the calendar
            for(var j = 0; j < 24-1; j++) {
              row = j + 6;
              
              //  Find row for Time End on the calendar
              if(calendarTimeValues[j][0].valueOf() == timeEnd.valueOf()) {
                row -= 1;
                break;        
              }               
              
              //  Set color 
              SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(row,col,1,1).getA1Notation()).setBackground(getColor(eventType)); 
            }   

            //  Store the cell range of the event as static data
            currEvent = SHEET_CALENDAR.getRange(SHEET_CALENDAR.getRange(6,col,(row + 1) - 6,1).getA1Notation());
            eventRanges.push(currEvent);                   
          }
        }
    }

    //  -------------------------------------------------------------------------------------
    //  Set index variables for looping through data    
    var index = 1;

    //  Merge all cells that make up each event
    eventRanges.forEach(function(value) { 
      var cell = SHEET_STATIC_DATA.getRange(index,1);
      cell.setValue(value.getA1Notation());
      SHEET_CALENDAR.getRange(value.getA1Notation()).merge();
      index++;
      });
}

//  -------------------------------------------------------------------------------------
//                                CALENDAR LAYOUT FUNCTIONS
//  -------------------------------------------------------------------------------------

//  -------------------------------------------------------------------------------------
//  Show columns of dates on the calendar

function showCols() { 
    //  Grab the list of dates
    var values = RANGE_CALENDAR_DATE.getValues();

    //  Show each date of the trip
    for (var i = 0; i < values[0].length - 1; i++) {
        SHEET_CALENDAR.showColumns(i + 1);
    }   
}

//  -------------------------------------------------------------------------------------
//  Hide all dates on the calendar

function hideCols() {
    //  Grab the list of dates
    var values = RANGE_CALENDAR_DATE.getValues();

    //  Hide each shown date on the calendar
    for (var i = 0; i < values[0].length - 1; i++) {   

        if (values[0][i] == "") {
            SHEET_CALENDAR.hideColumns(i + 2);
        }
    }
}

//  -------------------------------------------------------------------------------------
//  Merge the title cell of the calendar sheet to match the length of the trip (for clean UI)

function mergeCalendarTitle() {
    //  Unmerge the title cell
    RANGE_CALENDAR_TITLE.breakApart();

    //  Grab the trip length data
    var days =SHEET_TRIP_DETAILS.getRange('C5').getValue();
    var mergeRange = SHEET_CALENDAR.getRange(1, 1, 1, 1 + days);

    //  Merge the title cell
    mergeRange.mergeAcross();
}
