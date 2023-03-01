<?php
class SpreadSheet {
	function __construct (
			$name,
			$key,
			$tablerange,
			$timezone=null,
			$credentials=null,
			$datecolumn=1,
			$titlecolumn=2,
			$subtitlecolumn=null,
			$useheader=true,
			$combine=false,
			$allowall=true,
			$entries=null,
			$suffixentry=false,
			$casesensitive=false,
			$link=null,
			$group=true) {
		
		$this->name = $name;
		$this->key = $key;
		$this->tablerange = $tablerange;
		$this->timezone = $timezone;
		$this->credentials = $credentials;
		$this->datecolumn = $datecolumn;
		$this->titlecolumn = $titlecolumn;
		$this->subtitlecolumn = $subtitlecolumn;
		$this->useheader = $useheader;
		$this->combine = $combine;
		$this->allowall = $allowall;
		$this->entries = $entries;
		$this->suffixentry = $suffixentry;
		$this->casesensitive = $casesensitive;
		$this->link = $link;
		$this->group = $group;
	}
}

class CalendarEntry {
	function __construct ($title, $start, $length=1, $offset=0, $allday=false, $suffixentry=false) {
	 	$this->title = $title;
	 	$this->start = $start;		//time 00:00:00
	 	$this->length = $length;	//hours
	 	$this->allday = $allday;
	 	$this->offset = $offset;	//days
	 	$this->suffixentry = $suffixentry;
	}
}

class CalendarEvent {
	public $date = null;
	public $title = null;
	public $subtitle = null;
	public $containsitem = false;
	public $row = null;
	public $items = array();

	function __construct () {

	}
}

function selfURL($showparams = true, $protocol = null) {
	$s = empty($_SERVER["HTTPS"]) ? '' : ($_SERVER["HTTPS"] == "on") ? "s" : "";
	$protocol = (isset($protocol) ? $protocol : strleft(strtolower($_SERVER["SERVER_PROTOCOL"]), "/").$s);
	$port = ($_SERVER["SERVER_PORT"] == "80") ? "" : (":".$_SERVER["SERVER_PORT"]);
	$uri = $_SERVER['REQUEST_URI'];
	$pos = strrpos($_SERVER['REQUEST_URI'],'?');
	return $protocol."://".$_SERVER['SERVER_NAME'].$port.($showparams || $strrpos === false ? $uri : substr($uri,0,$pos));
}

function strleft($s1, $s2) {
	return substr($s1, 0, strpos($s1, $s2));
}

function isblank($var) {
	return (!isset($var) || $var == '');
}

function isdate($var) {
	//TODO: check for more?
	return $var != '' && $var != 0 && is_numeric($var);
}

function echotest($var) {
	global $Test;
	if ($Test) {
		echo $var;
	}
}

function serialtoepoch($serial, $fix = true) {
	$serialdate = floor($serial);
	$serialtime = $serial - $serialdate;
	//1 Jan 1900 = 1 in MS 31 Dec 1989 in Google
	//29 Jan 1900 = 60 in MS, account for excel incorrectly setting leap year in 1900 like lotus
	//See - http://support.microsoft.com/default.aspx?scid=kb;en-us;214019
	$serialdate = $serialdate - ($fix && $serialdate >= 60 ? 1 : 0);
	//http://www.cpearson.com/excel/datetime.htm
	$Jan1900 = strtotime('31 December 1899');
	$obj = new DateTime();
	$obj->setTimestamp(strtotime('31 December 1899')); //31 because Excel is from 0 January 1900
	$int = new DateInterval('P'.abs($serialdate).'D');
	$new = $serialdate > 0 ? $obj->add($int) : $obj->sub($int);
	return $new->getTimestamp();
}

function addical($iCal, $sheet, $event, $header, $Item, $Not, $Test=false) {
	if (isset($sheet->timezone)) {
		$current = date_default_timezone_get();
		date_default_timezone_set($sheet->timezone);
	}
	if (isset($event)) {
		if ((isblank($Item) && isblank($Not)) || ($event->containsitem && !isblank($Item)) ||  (!$event->containsitem && !isblank($Not))) {
			foreach ($sheet->entries as $entry) {
				$title = $event->title ? $event->title : '';
				$title .= isblank($entry->title) ? '' : (isblank($title) ? '' : ' ').$entry->title;
				if ($sheet->suffixentry) {
					foreach ($event->items as $item) {
						$title = $title . (isblank($title) ? '' : ' ') . $item['value'];
					}
				}
				$description = '';
				/*if ($sheet->group) {
					$final = array();
					foreach ($event->items as $item) {
						$found = false;
						foreach ($final as $key => $value) {
							if ($key == $item['key']) {
								$final[$key] = $value . ', ' . $item['value'];
								$found = true;
							}
						}
						if (!$found) {
							$final[$item['key']] = $item['value'];
						}
					}
					foreach ($final as $key => $value) {
						$description = $description . (isblank($description) ? '' : ', ') . (isblank($key) ? '' : $key . ': ') . $value;
					}
				} else {
					foreach ($event->items as $item) {
						//TODO: Combine same keys
						$description = $description . (isblank($description) ? '' : ', ') . (isblank($item['key']) ? '' : $item['key'] . ': ') . $item['value'];
					}
				}*/
				if ($entry->allday) {
					//$start = serialtoepoch($event->date + $entry->offset); //YYYY-mm-dd 00:00:00
					$start = $event->date;
					$end = 'allday';
				} else {
					$times = split(':',$entry->start);
					//echo date("Ymd H:i:s <br/>",serialtoepoch($event->date + $entry->offset))
					//$start = serialtoepoch($event->date + $entry->offset) + $times[2] + 60*($times[1] + 60*$times[0]); //TODO: Include start time
					$start = $event->date;
					$end = (int) ($start + 60*60*$entry->length);
				}
				$iCal->addEvent(
					array(), // Organizer
					(int) $start, // Start Time (timestamp; for an allday event the startdate has to start at YYYY-mm-dd 00:00:00)
					$end, // End Time (write 'allday' for an allday event instead of a timestamp)
					'', // Location
					0, // Transparancy (0 = OPAQUE | 1 = TRANSPARENT)
					array(), // Array with Strings
					$description, // Description
					$title, // Title
					1, // Class (0 = PRIVATE | 1 = PUBLIC | 2 = CONFIDENTIAL)
					array(), // Array (key = attendee name, value = e-mail, second value = role of the attendee [0 = CHAIR | 1 = REQ | 2 = OPT | 3 =NON])
					5, // Priority = 0-9
					0, // frequency: 0 = once, secoundly - yearly = 1-7
					10, // recurrency end: ('' = forever | integer = number of times | timestring = explicit date)
					2, // Interval for frequency (every 2,3,4 weeks...)
					array(), // Array with the number of the days the event accures (example: array(0,1,5) = Sunday, Monday, Friday
					0, // Startday of the Week ( 0 = Sunday - 6 = Saturday)
					'', // exeption dates: Array with timestamps of dates that should not be includes in the recurring event
					'',  // Sets the time in minutes an alarm appears before the event in the programm. no alarm if empty string or 0
					1, // Status of the event (0 = TENTATIVE, 1 = CONFIRMED, 2 = CANCELLED)
					($sheet->link == null ? 'https://docs.google.com/spreadsheets/d/' . $sheet->key : $sheet->link), // optional URL for that event
					'en', // Language of the Strings
					$sheet->key.'_'.$event->date.'_'.$event->row // Optional UID for this event
				);
			}
		}
	}
	if (isset($sheet->timezone)) {
		date_default_timezone_set($current);
	}
}

?>