<?php
error_reporting(E_ALL ^ (E_NOTICE | E_WARNING | E_DEPRECATED));
$Version = '1.4.2';
$URL = 'http://jameslow.com/2008/11/15/google-spreadsheet-to-ical/';

require_once 'lib/google-api-php-client/vendor/autoload.php';
require_once 'lib/ical/class.iCal.inc.php';
$iCal = (object) new iCal('', 0, ''); // (ProgrammID, Method (1 = Publish | 0 = Request), Download Directory)
require_once 'include.php';
require_once 'config.php';

function sheetlink($link, $text) {
	echo '<a href="'.$link.'">'.$text.'</a><br />';
}

if (isset($_REQUEST['help'])) {
	echo '<html><head>';
	echo '</head><body>';
	echo 'Google Spreadsheet to iCal ' . $Version . '<br />';
	echo '<a href="'.$URL.'">'.$URL.'</a><br />';
	echo '<br />';
	echo 'Avaliable Sheets:<br />';
	$masterlink = selfURL(false,'webcal');
	$mastertext = selfURL(false);
	foreach ($SHEETS as $sheet) {
		$sub = (substr($masterlink,strlen($masterlink)-1) == '/' ? '' : '/') . '?sheet=' . $sheet->name;
		$link = $masterlink . $sub;
		$text = $mastertext . $sub;
		sheetlink($link, $text);
	}
	echo 'Default Sheet '.$DEFAULT_SHEET.':<br />';
	if ($DEFAULT_SHEET != '') {
		sheetlink($masterlink, $mastertext);
	}
	echo 'All Sheets:<br />';
	if ($ALLOW_ALL) {
		$sub = (substr($masterlink,strlen($masterlink)-1) == '/' ? '' : '/') . '?sheet=' . '[ALL]';
		$link = $masterlink . $sub;
		$text = $mastertext . $sub;
		sheetlink($link, $text);
	}
	echo '</body></html>';
} else {
$Test = isset($_REQUEST['test']);

$client = new \Google_Client();
$client->setApplicationName('Google Sheets API');
$client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
$client->setAccessType('offline');

//Which SpreadSheet to use
$SheetName = $_REQUEST['sheet'];
if (isblank($SheetName)) {
	$SheetName = $DEFAULT_SHEET;
}
if (isblank($SheetName)) {
	if ($ALLOW_CUSTOM) {
		$Key = $_REQUEST['key'];
		if (isblank($Key)) {
			die('sheet or key must be sepecified in URL, a $DEFAULT_SHEET in the config.');
		}
		//$Sheet = new SpreadSheet(
	} else {
		die('sheet must be sepecified in URL, a $DEFAULT_SHEET in the config, or $ALLOW_CUSTOM=true.');
	}
}

//Query Item Name
$Item = $_REQUEST["item"]; //TODO: should we get from config?
$Not = $_REQUEST["not"];

if ($SheetName == '[ALL]') {
	if ($ALLOW_ALL) {
		$Sheets = $SHEETS;
	} else {
		die('$ALLOW_ALL must be set to true .');
	}
} else {
	$Sheets = array();
	$sheets = explode(",",$SheetName);
	foreach ($SHEETS as $row) {
		foreach ($sheets as $sheet) {
			if ($row->name == $sheet) {
				$Sheets[] = $row;
			}
		}
	}
}

foreach ($Sheets as $sheet) {
	if (!isblank($Item) || $sheet->allowall) {
		$Credentials = $_REQUEST["credentials"];
		if (isblank($Credentials)) {
			$Credentials = $CREDENTIALS_PATH;
			if(!isblank($sheet->credentials)) {
				$Credentials = $sheet->credentials;
			}
		}

		$client->setAuthConfig($Credentials);
		$service = new \Google_Service_Sheets($client);
		//$spreadsheet = $service->spreadsheets->get($sheet->key);
		//print_r($spreadsheet);
		$response = $service->spreadsheets_values->get($sheet->key, $sheet->tablerange);
		$rows = $response->getValues();
		if (!$sheet->entries) {
			$sheet->entries = array(
				new CalendarEntry(null, '00:00:00', 24, 0, true, false)
			);
		}
		//TODO: make sure all day events
		//TODO: make sure showing multiple events

		$lastevent = null;
		$event = null;
		$header = null;
		$relrow = 1;
		foreach ($rows as $row) {
			$relcol = 1;
			foreach ($row as $value) {
				if ($firsttime) {
					$firsttime = false;
				}
				if ($row == $firstrow && $sheet->useheader) {
					$header[] = $value;
				} else {
					if ($relcol == 1) {
						$lastevent = $event;
						$event = new CalendarEvent();
						$event->row = $relrow;
					}
					if ($relcol == $sheet->datecolumn) {
						$value = strtotime($value);
						//Check if we have a new date, if we do, echo last event, otherwise combine
						//Note this means we only add if we have a date
						if(isdate($value) && (!$sheet->combine || !$lastevent || $value != $lastevent->date)) {
							addical($iCal,$sheet,$lastevent,$header,$Item,$Not,$Test);
							$event->date = $value;
						} elseif(isset($lastevent)) {
							$lastevent->items = array_merge($lastevent->items,$event->items);
							$event = $lastevent;
						}
						$lastevent = null;
					} elseif (isset($sheet->titlecolumn) && $relcol == $sheet->titlecolumn && $value != '') {
						$event->title = $value;
					} elseif (isset($sheet->subtitlecolumn) && $relcol == $sheet->subtitlecolumn && $value != '') {
						$event->subtitle = $value;
					} else {
						if (!isblank($value)) {
							$checkvalue = ($sheet->casesensitive ? $value : strtolower($value));
							$checkitem = ($sheet->casesensitive ? $Item : strtolower($Item));
							$checknot = ($sheet->casesensitive ? $Not : strtolower($Not));
							if ((!isblank($Item) && $checkvalue == $checkitem) || (!isblank($Not) && $checkvalue == $checknot)) {
								$event->containsitem = true;
							}
							if ($sheet->useheader) {
								$key = $header[$relcol-1];
							} else {
								$key = '';
							}
							$event->items[] = array('key' => $key, 'value' => $value);
						}
					}
				}
				$relcol++;
			}
			$relrow++;
		}
		if (isdate($event->date)) {
			addical($iCal,$sheet,$event,$header,$Item,$Not,$Test);
		}
	}
}
if (count($Sheets) == 1) {
	$iCal->setName($Sheets[0]->name);
} else {
	$iCal->setName($DEFAULT_NAME);
}
if ($Test) {
	echo $iCal->countiCalObjects();
} else {
	$iCal->outputFile('ics');
}

} ?>