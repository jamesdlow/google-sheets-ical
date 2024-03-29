<?php

$GENERAL_USER=''; //Global username to use if not specified for a certain sheet, or as part of the URL
$GENERAL_PASSWORD=''; //Global password to use if not specified for a certain sheet, or as part of the URL
$ALLOW_CUSTOM=false; //Allow a spreadsheet defiinition to be specified in the URL, currently not implemented yet
$ALLOW_ALL=true; //Allow all spreadsheets to be queried into a single feed
$DEFAULT_SHEET='[ALL]'; //Should match one of the spreasheet names defined below, or be left blank, or = [ALL], allows the URL to be used without specifying ?sheet=sheetname
$DEFAULT_NAME='jameslow.com iCal test'; //Default name when query more than one sheet

/*
An entry is an assoicated event created in the iCal feed for each new line in the spreadsheet file
A single entry in a roster can generate multiple
$entries[] = new CalendarEntry(
	$title, //Title of the calendar event in iCal feed
	$start, //Start time as a 24 hour time, eg. 15:30
	$length, //Length of the event in hours, default is 1
	$offset, //Offset of the entry from the date in the spreasheet in days, default is 0 days
	$allday, //Create all day event, will ignore $length, default is false
	$suffixentry //Currently not used
);


$SHEETS[] = new SpreadSheet(
	$name, //Name for this sheet when quried through the URL
	$key, //Key for this spreadsheet, something like pgDM1_IcKknpu0qD6viidFg
	$tablerange, //Range to use for the table, this seems to work better with hardcoded range eg. A1:C15, rather than a named range
	$timezone, //Timezone for thie calendar, default is nothing
	$username, //Username to access this sheet
	$password, //Password to access this sheet
	$worksheet, //Worksheet index of the table to be used, eg. first spreadsheet is 1, default is 1
	$datecolumn, //Date column relative to range, eg. first column is 1, default is 1
	$useheader, //Get item subheaders from first row, default is true
	$subtitlecolumn, //Collumn to get a subtitle from, currently not used
	$allowall, //Allow showing of all roster entries, not just those for a single queried item, default is true
	$entries, //Array of CalendarEntry, see above example
	$suffixentry, //Suffix entry title for data found in column, this allows the title or a suffix to the title to be specified in the spreadsheet, default is false
	$casesensitive, //Make item query case sensitive, default is false
	$link, //Link to be shown in this calendar event, default is a link back to the spreadsheet
	$group //Group column data if each date spans multiple rows, default is true
);
*/

$SHEETS[] = new SpreadSheet('TestSheet','pgDM1_IcKknpu0qD6viidFg','B3:F29','Asia/Hong_Kong',null,null,1,1,true,null,true,array(new CalendarEntry('Google iCal test', '16:00', 5)));
$entries[] = new CalendarEntry('Multi Entry Test 1', '10:00', 5, -1);
$entries[] = new CalendarEntry('Multi Entry Test 2', '16:00', 5);
$SHEETS[] = new SpreadSheet('MultiEntry','pgDM1_IcKknpu0qD6viidFg','B3:F29','Asia/Hong_Kong',null,null,1,1,true,null,true,$entries);

?>