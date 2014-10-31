<?php
error_reporting(E_ALL ^ (E_NOTICE | E_WARNING | E_DEPRECATED));
$Version = '1.4.2';
$URL = 'http://jameslow.com/2008/11/15/google-spreadsheet-to-ical/';
$clientLibraryPath = '.'.substr(__DIR__,strlen(getcwd()));
$oldPath = set_include_path(get_include_path() . PATH_SEPARATOR . $clientLibraryPath);

require_once 'Zend/Loader.php';
Zend_Loader::loadClass('Zend_Gdata');
Zend_Loader::loadClass('Zend_Gdata_ClientLogin');
Zend_Loader::loadClass('Zend_Gdata_Spreadsheets');
Zend_Loader::loadClass('Zend_Gdata_App_AuthException');
Zend_Loader::loadClass('Zend_Http_Client');
require_once 'ical/class.iCal.inc.php';
$iCal = (object) new iCal('', 0, ''); // (ProgrammID, Method (1 = Publish | 0 = Request), Download Directory)
require_once 'include.php';
require_once 'config.php';

function sheetlink($link, $text) {
	echo '<a href="'.$link.'">'.$text.'</a><br />';
}

if (isset($_REQUEST['help'])) {
	echo '<html><head>';
	echo '<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
try {
_uacct = "UA-10975515-5";
urchinTracker();
} catch(err) {}</script>';
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
$Item = $_REQUEST["item"];
$Not = $_REQUEST["not"];

if ($SheetName == '[ALL]') {
	if ($ALLOW_ALL) {
		$Sheets = $SHEETS;
	} else {
		die('$ALLOW_ALL must be set to true .');
	}
} else {
	$Sheets = array();
	$sheets = split(",",$SheetName);
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
		//Username and Password
		$User = $_REQUEST["user"];
		$Password = $_REQUEST["password"];
		if (isblank($User)) {
			$User = $GENERAL_USER;
			$Password = $GENERAL_PASSWORD;
			if(!isblank($sheet->username)) {
				$User = $sheet->username;
				$Password = $sheet->password;
			}
			if(!isset($User)) {
				$User = '';
				$Password = '';
			}
		}

		$service = Zend_Gdata_Spreadsheets::AUTH_SERVICE_NAME;
		$client = Zend_Gdata_ClientLogin::getHttpClient($User, $Password, $service);
		$spreadsheetService = new Zend_Gdata_Spreadsheets($client);
		$query = new Zend_Gdata_Spreadsheets_CellQuery();
		$query->setSpreadsheetKey($sheet->key);
		$query->setWorksheetId($sheet->worksheet);
		$query->setRange($sheet->tablerange);
		$cellFeed = $spreadsheetService->getCellFeed($query);

		$firsttime=true;
		$lastrow=0;
		$date='';
		$lastevent=null;
		$event=null;
		$header = null;

		foreach ($cellFeed as $cellEntry) {
			$row = $cellEntry->cell->getRow();
			$col = $cellEntry->cell->getColumn();
			$value = $cellEntry->cell->getText();
			if ($Test) {
				//echo $value;
			}
			if($firsttime) {
				$rowoffset = $row-1;
				$coloffset = $col-1;
				$firstrow = $row;
				$firstcol = $col;
				$lastcol = $col+$cellFeed->getColumnCount()-1;
				$lastrow = $row+$cellFeed->getRowCount()-1;
				$firsttime=false;
			}
			$relcol = $col - $coloffset;
			$relrow = $row - $rowoffset;
			if ($row == $firstrow && $sheet->useheader) {
				$header[] = $value;
			} else {
				if ($Test) {
					//print_r($header);
				}
				if ($col == $firstcol) {
					$lastevent = $event;
					$event = new CalendarEvent();
				}
				if ($relcol == $sheet->datecolumn) {
					$value = $cellEntry->cell->getNumericValue();
					//Check if we have a new date, if we do, echo last event, otherwise combine
					if(isdate($value) && $value != $lastevent->date) {
						addical($iCal,$sheet,$lastevent,$header,$Item,$Not,$Test);
						$event->date = $value;
					} elseif(isset($lastevent)) {
						$lastevent->items = array_merge($lastevent->items,$event->items);
						$event = $lastevent;
					}
					$lastevent = null;
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
		}
	addical($iCal,$sheet,$event,$header,$Item,$Not,$Test);
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