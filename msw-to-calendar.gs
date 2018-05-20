function main() { 
  fillCalendar(createCalendar(), extractRow());
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Fill calendar for selected person', 'main');
  menu.addToUi();
}

function fillCalendar(calendar, row) {
  var beginningOfTheMonth = getBeginningOfTheMonth();
  var month = beginningOfTheMonth.getMonth();
  var year = beginningOfTheMonth.getFullYear();

  // TODO: check if number of days is correct
  for (var i = 0; i < row.length; i++) {
    var period = hours[row[i]];
    if (period) {
      var fromDate = new Date(year, month, i+1, period[0], period[1]);
      var toDate = new Date(year, month, i+1, period[2], period[3]);
      // holiday
      // TODO: extract edge cases to separate functions
      if (period[0] == period[2] && period[1] == period[3]) {
        calendar.createAllDayEvent("Urlop", fromDate)
      } else  {
        // night shift
        if (period[0] > period[2]) toDate.setDate(toDate.getDate() + 1);
        calendar.createEvent(row[i], fromDate, toDate);
      }
    }
  }
}

function createCalendar() {
  var calendar;
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length == 1) {
    calendar = calendars[0];
  } else {
    calendar = CalendarApp.createCalendar(calendarName);
  }
  // TODO: extract to a separate function (delete all events in the month)
  var beginningOfTheMonth = getBeginningOfTheMonth();
  var endOfTheMonth = new Date(beginningOfTheMonth.getFullYear(), beginningOfTheMonth.getMonth() + 1, 0);
  var events = calendar.getEvents(beginningOfTheMonth, endOfTheMonth);
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
  return calendar;
}

function getBeginningOfTheMonth() {
  var now = new Date();
  var month = getMonth();
  var year = now.getFullYear();
  return new Date(year, month, 1);
}

function extractRow() {
  var selectedPersonRow = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  var row = [];
  var column = 2;
  var day = SpreadsheetApp.getActiveSheet().getRange(2, column);
  while (day.isBlank() || !isNaN(parseInt(day.getValue()))) {
    if (!day.isBlank()) {
      row.push(SpreadsheetApp.getActiveSheet().getRange(selectedPersonRow, column).getValue());
    }
    column++;
    day = SpreadsheetApp.getActiveSheet().getRange(2, column);
  }
  Logger.log(row);
  return row;
}

function getMonth() {
  var monthString = SpreadsheetApp.getActiveSheet().getRange('A1').getDisplayValue();
  return monthNumber[monthString.toLowerCase().escapeDiacritics()];
}  

String.prototype.escapeDiacritics = function()
{
    return this.replace(/ą/g, 'a').replace(/Ą/g, 'A')
        .replace(/ć/g, 'c').replace(/Ć/g, 'C')
        .replace(/ę/g, 'e').replace(/Ę/g, 'E')
        .replace(/ł/g, 'l').replace(/Ł/g, 'L')
        .replace(/ń/g, 'n').replace(/Ń/g, 'N')
        .replace(/ó/g, 'o').replace(/Ó/g, 'O')
        .replace(/ś/g, 's').replace(/Ś/g, 'S')
        .replace(/ż/g, 'z').replace(/Ż/g, 'Z')
        .replace(/ź/g, 'z').replace(/Ź/g, 'Z');
}

var monthNumber = new Object();
monthNumber['styczen'] = 0;
monthNumber['luty'] = 1;
monthNumber['marzec'] = 2;
monthNumber['kwiecien'] = 3;
monthNumber['maj'] = 4;
monthNumber['czerwiec'] = 5;
monthNumber['lipiec'] = 6;
monthNumber['sierpien'] = 7;
monthNumber['wrzesien'] = 8;
monthNumber['pazdziernik'] = 9;
monthNumber['listopad'] = 10;
monthNumber['grudzien'] = 11;

var calendarName = 'MSW-praca'
var eventName = 'Praca'

var hours = new Object();
hours['R0'] = [8,00,14,00];
hours['R1'] = [9,00,15,00];
hours['DK'] = [7,00,19,00];
hours['D'] = [8,00,20,00];
hours['P'] = [9,00,16,30];
hours['R'] = [7,30,15,00];
hours['N'] = [19,00,7,00];
hours['r'] = [6,30,14,30];
hours['pp'] = [7,00,14,00];
hours['U'] = [0,00,0,00];
hours['A'] = [8,00,15,30];
hours['Ssz'] = [7,30,15,00];
hours['Rsz'] = [7,30,15,00];
hours['Dsz'] = [7,30,19,30];
hours['r/Ssz'] = [7,00,19,00];
hours['pp1/Ssz'] = [7,00,19,00];
hours['pp2/Ssz'] = [7,00,19,00];
