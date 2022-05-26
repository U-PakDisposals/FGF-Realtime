<?php
#region
require 'C:\xampp\htdocs\trucknotes\Composer\vendor\autoload.php';
require 'C:\xampp\htdocs\trucknotes\Composer\vendor\PHPMailer-master\PHPMailer-master\src\PHPMailer.php';
require 'C:\xampp\htdocs\trucknotes\Composer\vendor\PHPMailer-master\PHPMailer-master\src\Exception.php';
require 'C:\xampp\htdocs\trucknotes\Composer\vendor\PHPMailer-master\PHPMailer-master\src\SMTP.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as writer;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as reader;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

#endregion
#region Excel Rows (Address)
$sNameRow = 'A';
$sAddressRow = 'B';
$sDateRow = 'C';
$sTimeRow = 'D'; // NEW
$sDescRow = 'E';
$sWoNumRow = 'F';
$sVehRow = 'G';
$sRouteRow = 'H';
$sCountRow = 'I';
$sDoneRow = 'J';
$sTypeRow = 'K';
$aCodeRow = 'L';
#endregion
#region Excel Rows (Main)
$mWoNumRow = 'A';
$mDateRow = 'B';
$mTimeRow = 'C'; // NEW
$mVehRow = 'D';
$mCountRow = 'E';
#endregion

class Service
{
	#region variables
	private $custName;
	private $custAddress;
	private $serviceDateTime;
	private $serviceDate;
	private $serviceTime;
	private $serviceDescription;
	private $workorderId;
	private $vehicle;
	private $route;
	private $serviceCode;
	private $activityCode;
	#endregion
			//new Service($custName, $custAddress, $serviceDateTime, $serviceDescription, $workorderId, $vehicle, $route, $serviceCode)
	function __construct($custName, $custAddress, $serviceDateTime, $serviceDescription, $workorderId, $vehicle, $route, $serviceCode, $activityCode)
	{
		$this->setCustName($custName);
		$this->setCustAddress($custAddress);
		$this->setserviceDateTime($serviceDateTime);
		$this->setServiceDescription($serviceDescription);
		$this->setWorkorderId($workorderId);
		$this->setVehicle($vehicle);
		$this->setRoute($route);
		$this->setServiceCode($serviceCode);
		$this->setActivityCode($activityCode);
	}
	#region setters and getters
	public function setCustName($custName) 									{$this->custName = $custName;}
	public function getCustName() 											{return $this->custName;}

	public function setCustAddress($custAddress) 							{$this->custAddress = $custAddress;}
	public function getCustAddress() 										{return $this->custAddress;}

	public function setServiceDateTime($serviceDateTime)							{
		$this->serviceDate = explode("T", $serviceDateTime)[0];
		$this->serviceTime = explode("T", $serviceDateTime)[1];
	}
	public function getServiceDate()										{return $this->serviceDate;}
	public function getServiceTime()										{return $this->serviceTime;}

	public function setServiceDescription($serviceDescription)				{$this->serviceDescription = $serviceDescription;}
	public function getServiceDescription()									{return $this->serviceDescription;}

	public function setWorkorderId($workorderId)							{$this->workorderId = $workorderId;}
	public function getWorkorderId()										{return $this->workorderId;}

	public function setVehicle($vehicle)									{$this->vehicle = $vehicle;}
	public function getVehicle()											{return $this->vehicle;}

	public function setRoute($route)										{$this->route = $route;}
	public function getRoute()												{return $this->route;}

	public function setServiceCode($serviceCode)							
	{
		switch ($serviceCode)
		{
			case 'GE':
				$this->serviceCode = 'DUMP & RETURN';
				break;
			case 'BE':
				$this->serviceCode = 'EXCHANGE';
				break;
			case 'CE':
				if ($this->custAddress == '100 Locke St' || $this->custAddress == '1295 Ormont')
				{
					$this->serviceCode = 'EXCHANGE';
					break;
				} else
				{
					$this->serviceCode = 'DUMP & RETURN';
					break;
				}
			case 'CD':
				$this->serviceCode = 'EXCHANGE';
				break;
			case 'GD':
				$this->serviceCode = 'DUMP & RETURN';
				break;
			case 'ET':
				$this->serviceCode = 'EXCHANGE';
				break;
			case 'CX':
				$this->serviceCode = 'FREE SERVICE - DUMP & RETURN';
				break;
			default:
					$this->serviceCode = 'SERVICE CODE NOT SET, ' . $serviceCode;
					break;
		}
	}
	public function getServiceCode()										{return $this->serviceCode;}

	public function setActivityCode($activityCode)
	{
		switch ($activityCode)
		{
			case 'B':
				$this->activityCode = 'BLOCKED';
				break;
			case 'E':
				$this->activityCode = 'EMPTY';
				break;
			default:
				$this->activityCode = 'Activity Code: ' . $activityCode;
				break;
		}
	}
	public function getActivityCode()										{return $this->activityCode;}
	#endregion
}
$excel = IOFactory::load('C:\xampp\htdocs\trucknotes\FGF\FGFRealtime-ENHANCED.xlsx');
$serviceArray = array();
error_reporting(1);
date_default_timezone_set('America/Toronto');
$startDate = date('Y-m-d H:i', time() - 300); // 1800 = half an hour, 300 = 5 minutes // $startDate = '2022-05-06T13:00';
$startDate = str_replace(' ', 'T', $startDate);
$startDate .= ':00.000Z';
$endDate = date('Y-m-d H:i', time()); // $endDate = '2022-05-06T13:30';
$endDate = str_replace(' ', 'T', $endDate);
$endDate .= ':00.000Z';
#region setting up connection
$searchCriteria = "{\"startDate\": \"" . $startDate . "\",\"endDate\": \"" . $endDate . "\"}";
echo $searchCriteria;

$division = '-488398130';
$apiKey = 'AIzaSyAfdcuAmGCApthHPx18Vx3ohoM4rgXWcX0';
$url = 'https://upak.fleetmindcloud.com/FleetmindAPI/v2/division/' . $division . '/report/services';
$headers = array();
$headers[] = 'Content-Type: application/json';
$headers[] = 'Accept: application/json';
$headers[] = 'X-Apikey: ' . $apiKey;

$ch = curl_init();
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_URL, $url);
curl_setopt($ch, CURLOPT_POSTFIELDS, $searchCriteria);
curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);

$jsonExec = curl_exec($ch);
if(curl_errno($ch))
{
	echo 'Error:' . curl_error($ch);
}

curl_close($ch);

$json = json_decode($jsonExec, true); 
$length = count($json['Data'], 0);
echo $length . ' service calls found<br><br>';
#endregion
if ($length > 0)
{
    for ($i = 0; $i < $length; $i++) // create new service for each service
    {
        if (strpos($json['Data'][$i]['customerName'], 'FGF') !== FALSE)
        {
			// echo $json['Data'][$i]['customerName'] . '<br>';
			$custName = $json['Data'][$i]['customerName'];
			$custAddress = $json['Data'][$i]['customerAddress'];
			$serviceDateTime = $json['Data'][$i]['Details']['serviceTime'];
			$serviceDescription = $json['Data'][$i]['Details']['serviceDescription'];
			$expectedLat = $json['Data'][$i]['Details']['expectedLatitude'];
			$expectedLong = $json['Data'][$i]['Details']['expectedLongitude'];
			$servicedLat = $json['Data'][$i]['Details']['servicedLatitude'];
			$servicedLong = $json['Data'][$i]['Details']['servicedLongitude'];
			$workorderId = $json['Data'][$i]['Details']['workOrderId'];
			$vehicle = $json['Data'][$i]['ConfirmedVehicle']['VehicleName'];
			$route = $json['Data'][$i]['ConfirmedVehicle']['Route'];
			$serviceCode = $json['Data'][$i]['Details']['serviceCode'];
			$activityCode = $json['Data'][$i]['Details']['activityCode'];
			array_push($serviceArray, new Service($custName, $custAddress, $serviceDateTime, $serviceDescription, $workorderId, $vehicle, $route, $serviceCode, $activityCode));
        }
    }
	foreach($serviceArray as $bin) // inserts data in Excel
	{
		insertToExcel($excel, $bin);
	}
	$sheetCount = $excel->getSheetCount();
	$sheetDone = $excel->getSheetByName('Done');
	for ($i = 1; $i < $sheetCount; $i++) // for each spreadsheet in the excel sheet
	{
		$mainSheet = $excel->getSheet($i);
		for ($h = 50; $h >= 2; $h--)
		{
			if ($mainSheet->getCell($sCountRow . $h)->getValue() == '2' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && $mainSheet->getCell($sTypeRow . $h)->getValue() == 'EXCHANGE')
			{
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, 'REPLACING BIN');
			}
			elseif ($mainSheet->getCell($sCountRow . $h)->getValue() == '2' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && $mainSheet->getCell($sTypeRow . $h)->getValue() == 'DUMP & RETURN')
			{
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, 'PICKING UP BIN, DROPPING OFF LATER');
			}
			elseif ($mainSheet->getCell($sCountRow . $h)->getValue() == '2' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && $mainSheet->getCell($sTypeRow . $h)->getValue() == 'FREE SERVICE - DUMP & RETURN')
			{
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, 'FREE SERVICE - PICKING UP BIN, DROPPING OFF LATER.');
			}
			elseif ($mainSheet->getCell($sCountRow . $h)->getValue() == '5' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && $mainSheet->getCell($sTypeRow . $h)->getValue() == 'FREE SERVICE - DUMP & RETURN')
			{ 
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, 'FREE SERVICE - RETURNING BIN');
			}
			elseif ($mainSheet->getCell($sCountRow . $h)->getValue() == '5' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && $mainSheet->getCell($sTypeRow . $h)->getValue() == 'DUMP & RETURN')
			{ 
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, 'RETURNING BIN');
			}
			elseif ($mainSheet->getCell($sCountRow . $h)->getValue() == '2' && $mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && strpos($mainSheet->getCell($sTypeRow . $h)->getValue(), 'SERVICE CODE NOT SET') !== false)
			{ 
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, $bin->getServiceCode());
			}
			if ($mainSheet->getCell($sDoneRow . $h)->getValue() != 'DONE' && ($mainSheet->getCell($aCodeRow . $h)->getValue() == 'BLOCKED' || $mainSheet->getCell($aCodeRow . $h)->getValue() == 'EMPTY'))
			{
				moveToDone($sheetDone, $mainSheet, $h, $excel);
				$binService = excelToService($mainSheet, $h);
				sendEmail($binService, $bin->getActivityCode());
			}
		}
	}
}
function insertToExcel($excel, $bin)
{
	global $sNameRow;
	global $sAddressRow;
	global $sDateRow;
	global $sTimeRow;
	global $sDescRow;
	global $sWoNumRow;
	global $sVehRow;
	global $sRouteRow;
	global $sTypeRow;
	global $sCountRow;
	global $aCodeRow;

	$sheetArray = $excel->getSheetNames();
	if (in_array($bin->getCustAddress(), $sheetArray))
	{
		// echo $bin->getCustAddress() . ' sheet found<br>';
	}
	else // if sheet not found, create new sheet
	{
		echo $bin->getCustAddress() . ' sheet not found<br>';
		$excel->addSheet(new Worksheet($excel, $bin->getCustAddress()));
	}
	$counter = 1;
	$activeSheet = $excel->getSheetByName($bin->getCustAddress());
	$activeSheet->insertNewRowBefore(2, 1); // inserts "1" new row before row "2"
	$lastSheetRow = $activeSheet->getHighestRow();
	$activeSheet->setCellValue($sNameRow . '2', $bin->getCustName());
	$activeSheet->setCellValue($sAddressRow . '2', $bin->getCustAddress());
	$activeSheet->setCellValue($sDateRow . '2', $bin->getServiceDate());
	$activeSheet->setCellValue($sTimeRow . '2', $bin->getServiceTime());
	$activeSheet->setCellValue($sDescRow . '2', $bin->getServiceDescription());
	$activeSheet->setCellValue($sWoNumRow . '2', $bin->getWorkorderId());
	$activeSheet->setCellValue($sVehRow . '2', $bin->getVehicle());
	$activeSheet->setCellValue($sRouteRow . '2', $bin->getRoute());
	$activeSheet->setCellValue($sTypeRow . '2', $bin->getServiceCode());
	$activeSheet->setCellValue($aCodeRow . '2', $bin->getActivityCode());
	for ($k = $lastSheetRow; $k >= 2; $k--)
	{
		// echo 'checking...<br>';
		// echo $activeSheet->getCell('I' . $k)->getValue() . ' ' . $bin->getWorkorderId() . '<br>';
		if ($activeSheet->getCell($sWoNumRow . $k)->getValue() == $bin->getWorkorderId() && $activeSheet->getCell($sVehRow . $k)->getValue() == $bin->getVehicle())
		{
			$activeSheet->setCellValue($sCountRow . $k, $counter++);
		}
	}

	save($excel);
}
function excelToService($mainSheet, $rowNum)
{ // new Service($custName, $custAddress, $serviceDateTime, $serviceDescription, $workorderId, $vehicle, $route, $serviceCode)
	global $sNameRow;
	global $sAddressRow;
	global $sDateRow;
	global $sTimeRow;
	global $sDescRow;
	global $sWoNumRow;
	global $sVehRow;
	global $sRouteRow;
	global $sTypeRow;
	global $aCodeRow;

	$bin = new Service($mainSheet->getCell($sNameRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sAddressRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sDateRow . $rowNum)->getValue() . 'T' . $mainSheet->getCell($sTimeRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sDescRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sWoNumRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sVehRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sRouteRow . $rowNum)->getValue(),
					   $mainSheet->getCell($sTypeRow . $rowNum)->getValue(),
					   $mainSheet->getCell($aCodeRow . $rowNum)->getValue());
	return $bin;
}
function moveToDone($sheetDone, $mainSheet, $rowNum, $excel)
{
	global $mWoNumRow;
	global $mDateRow;
	global $mTimeRow;
	global $mVehRow;
	global $mCountRow;
	global $sDoneRow;
	global $sWoNumRow;
	global $sDateRow;
	global $sTimeRow;
	global $sVehRow;
	global $sCountRow;

	$sheetDone->insertNewRowBefore(2, 1);
	$sheetDone->setCellValue($mWoNumRow . '2', $mainSheet->getCell($sWoNumRow . $rowNum)->getValue());
	$sheetDone->setCellValue($mDateRow . '2', $mainSheet->getCell($sDateRow . $rowNum)->getValue());
	$sheetDone->setCellValue($mTimeRow . '2', $mainSheet->getCell($sTimeRow . $rowNum)->getValue());
	$sheetDone->setCellValue($mVehRow . '2', $mainSheet->getCell($sVehRow . $rowNum)->getValue());
	$sheetDone->setCellValue($mCountRow . '2', $mainSheet->getCell($sCountRow . $rowNum)->getValue());
	$mainSheet->setCellValue($sDoneRow . $rowNum, 'DONE');
	save($excel);
}
function save($excel)
{
	$writer = new Writer($excel);
	$writer->save('./FGFRealtime-ENHANCED.xlsx');
}
function sendEmail($bin, $pickReturn)
{
	$mail = new PHPMailer(true);
	$barmac590 = ['Albin@fgfbrands.com', 'Larry.Johnson@fgfbrands.com', 'Rocky.Caliw-Caliw@fgfbrands.com', 'ian.tabiolo@fgfbrands.com', 'Agosto.Lacwasan@fgfbrands.com', 'Ibrahim.Aburaneh@fgfbrands.com'];
	$creditstone777 = ['vikash.madras@fgfbrands.com', 'Katie.Keary@fgfbrands.com', 'carlo.del-rosario@fgfbrands.com', 'sanseeyan.thanabalasingam@fgfbrands.com', 'rodni.arcayera@fgfbrands.com', 'mohamed.abdulle@fgfbrands.com', 'jocid.palos@fgfbrands.com', 'Don.martin@fgfbrands.com', 'Rowena@fgfbrands.com'];
	$fenmar200 = ['vikash.madras@fgfbrands.com', 'Carlos.Tojanci@fgfbrands.com', 'elmer.palapus@fgfbrands.com', 'Abdullahi.Adan@fgfbrands.com', 'samuel.dolojol@fgfbrands.com', 'noel.agcaoili@fgfbrands.com', 'edgar.flores@fgfbrands.com', 'Muluberhan.Gobeze@fgfbrands.com', 'Ibrahim.Aburaneh@fgfbrands.com'];
	$locke100 = ['Albin@fgfbrands.com',	'Mahram.Adelyar@fgfbrands.com', 'Jesus.Julian@fgfbrands.com', 'Fernando.Maglaway@fgfbrands.com', 'Bukola.Faturoti@fgfbrands.com', 'Lucky.Osaretin@fgfbrands.com', 'John.Cervantes@fgfbrands.com', 'Jerry.Cabanlig@fgfbrands.com', 'Katie.Keary@fgfbrands.com'];
	$ormont1235 = ['Albin@fgfbrands.com', 'Vasisht.Dial@fgfbrands.com', 'mario.yangat@fgfbrands.com', 'Angelito.Rafin@fgfbrands.com', 'william.corpuz@fgfbrands.com', 'Michael.Musili@fgfbrands.com', 'Ibraahim.Aadan@fgfbrands.com', 'Ronilo.Galagaran@fgfbrands.com', 'Robin.Balbaloza@fgfbrands.com', 'Erwin.Santiago@fgfbrands.com', 'Katie.Keary@fgfbrands.com'];
	$ormont1295 = ['Albin@fgfbrands.com', 'pedro.echinique@fgfbrands.com', 'Eduardo.Padilla@fgfbrands.com', 'Leo.Permo@fgfbrands.com', 'Hamedkier.Zemicael@fgfbrands.com', 'Edwin.Valenzuela@fgfbrands.com', 'Katie.Keary@fgfbrands.com'];
	$woodslea10 = ['vikash.madras@fgfbrands.com', 'vikas.singh@fgfbrands.com', 'cyril.belisario@fgfbrands.com', 'mohamed.ali@fgfbrands.com', 'kwame.richards@fgfbrands.com', 'joel.asigbetsey@fgfbrands.com', 'mohamed.abubakar@fgfbrands.com', 'julie.espada@fgfbrands.com', 'hajji.villanueva@fgfbrands.com', 'Ibrahim.Aburaneh@fgfbrands.com'];
	

	try 
	{
		$mail->SMTPDebug = SMTP::DEBUG_SERVER;
		$mail->isSMTP();
		$mail->Host = 'smtp.office365.com';
		$mail->SMTPAuth = true;
		$mail->Username = 'belalkaoukji@upak.net';
		$mail->Password = 'fla+Door83';
		$mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
		$mail->Port = 587;
	
		$mail->setFrom('belalkaoukji@upak.net', 'FGF Bin');

		if (strpos($pickReturn, "SERVICE CODE NOT SET") !== false) {
			$mail->Body .= 'This Address Is not setup for FGF Emails.<br>';
		} else
		{
			switch ($bin->getCustAddress())
			{
				case '590 BARMAC DR':
					foreach($barmac590 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '777 CREDITSTONE RD':
					foreach($creditstone777 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '200 FENMAR DR':
					foreach($fenmar200 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '100 LOCKE ST':
					foreach($locke100 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '1235 ORMONT DR':
					foreach($ormont1235 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '1295 ORMONT DR':
					foreach($ormont1295 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				case '10 WOODSLEA RD':
					foreach($woodslea10 as $email)
					{
						$mail->addAddress($email);
					}
					break;
				default:
					$mail->Body .= 'This Address Is not setup for FGF Emails.<br>';
					break;
			}
		}

		$mail->addAddress('MichaelMcMillan@upak.net');
		$mail->addAddress('Dispatch@upak.net');
		$mail->addAddress('belalkaoukji@upak.net');
	
		$mail->isHTML(true);
		if (strpos($pickReturn, "BLOCKED") !== false || strpos($pickReturn, "EMPTY") !== false)
		{
			$mail->Subject = $pickReturn . '- FGF Bin Service';
			$mail->Body .= 'Please disregard recent email with Workorder ID: "' . $bin->getWorkorderId() . '"; bin is ';
		}
		else
		{
			$mail->Subject = 'FGF Bin Service';
		}
		if ($pickReturn == 'Activity Code: N OUT')
		{
			$mail->Subject = $pickReturn . '- FGF Bin Service';
			$pickReturn = 'BIN EMPTY.';
		}
		$mail->Body .= $pickReturn . '<br>';
		$mail->Body .= '<br>';
		$mail->Body .= 'Bin Address: ' . $bin->getCustAddress() . '<br>';
		$mail->Body .= 'Service Date: ' . $bin->getServiceDate() . '<br>';
		$mail->Body .= 'Service Time: ' . $bin->getServiceTime() . '<br>';
		$mail->Body .= 'Service Description: ' . $bin->getServiceDescription() . '<br>';
		$mail->Body .= 'Workorder ID: ' . $bin->getWorkorderId() . '<br>';
		$mail->send();
		echo 'Message has been sent';
		} catch (Exception $e)
		{
			echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
		}
}
?>