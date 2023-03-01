<?php
$CREDENTIALS_PATH = ''; //Credentials as per: https://www.nidup.io/blog/manipulate-google-sheets-in-php-with-api
$ALLOW_CUSTOM     = false; //Allow a spreadsheet defiinition to be specified in the URL, currently not implemented yet
$ALLOW_ALL        = true; //Allow all spreadsheets to be queried into a single feed
$DEFAULT_SHEET    = '[ALL]'; //Should match one of the spreasheet names defined below, or be left blank, or = [ALL], allows the URL to be used without specifying ?sheet=sheetname
$DEFAULT_NAME     = 'jameslow.com iCal test'; //Default name when query more than one sheet

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
	$credentials, //Credentials path
	$datecolumn, //Date column relative to range, eg. first column is 1, default is 1
	$titlecolumn, //Title column if being read from spreadshseet, use null to get form entries
	$subtitlecolumn, //Column to get a subtitle from, can be null
	$useheader, //Get item subheaders from first row, default is true
	$combine, //Get item subheaders from first row, default is true
	$allowall, //Allow showing of all roster entries, not just those for a single queried item, default is true
	$entries, //Array of CalendarEntry, see above example
	$suffixentry, //Suffix entry title for data found in column, this allows the title or a suffix to the title to be specified in the spreadsheet, default is false
	$casesensitive, //Make item query case sensitive, default is false
	$link, //Link to be shown in this calendar event, default is a link back to the spreadsheet
	$group //Group column data if each date spans multiple rows, default is true
);
*/

$SHEETS[] = new SpreadSheet('TestSheet','pgDM1_IcKknpu0qD6viidFg','B3:F29','UTC',null,1,null,null,true,true,true,array(new CalendarEntry('Google iCal test', '16:00', 5)));
$entries[] = new CalendarEntry('Multi Entry Test 1', '10:00', 5, -1);
$entries[] = new CalendarEntry('Multi Entry Test 2', '16:00', 5);
$SHEETS[] = new SpreadSheet('MultiEntry','pgDM1_IcKknpu0qD6viidFg','B3:F29','UTC',null,1,null,null,true,true,true,$entries);

?>