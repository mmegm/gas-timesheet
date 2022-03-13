function Terminesammeln(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
 SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
   SpreadsheetApp.getActiveSheet()
   
// Kalender verknüpfen
var mycal = ss.getRange("B8").getDisplayValue();
var cal = CalendarApp.getCalendarById(mycal);


var now = Utilities.formatDate(new Date(), "GMT+01:00", "dd-MM-YY HH:mm:ss");
  // Variablen aus Zellen auslesen
  var suchwort = ss.getRange("B9").getDisplayValue();
  var startdatum  = ss.getRange("B6").getValue();
  var enddatum = ss.getRange("B7").getValue();

  
// Termine auslesen
var events = cal.getEvents(new Date(startdatum), new Date(enddatum), {search: suchwort});

 var sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[1]);

  // Blatt leeren
sheet.clearContents();  

// Tabelle schreiben
// ohne "Kalendername", 
var header = [["Titel", "" , "Beschreibung", "Ort", "Datum", "Enddatum", "Dauer", "Sichtbarkeit", "Erstellt", "Geändert", "Rechte", "Von", "ganzt", "wiederh."]]
var range = sheet.getRange(1,1,1,14);
range.setValues(header);

  
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
for (var i=0;i<events.length;i++) {
var row=i+2;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
// ohne mycal,
var details=[[events[i].getTitle(), " ", events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
var range=sheet.getRange(row,1,1,14);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');

  ss.getSheets()[1].setName("Stundenliste - " + suchwort + " - von " + now);
}
}
function onOpen() {

  
    var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
  
}

// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event

// Georg Mastritsch, 19.01.17
