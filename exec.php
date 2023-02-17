<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST');
header("Access-Control-Allow-Headers: X-Requested-With");

//get input data
$data_input = json_decode(file_get_contents('php://input'), true);

try{
// write to vbs file
    //$file_addr = 'C:\xampp\htdocs\online_elink\processed_files\rbg_master16312720032604.xlsm.vbs';
	$file_addr = 'C:\xampp\htdocs\online_elink\processed_files\tja_master16312838168408.xlsm.vbs';
	$output = null;
	$retval = null;

	//$output = shell_exec("$file_addr 2>&1");
	//var_dump($output);

	exec($file_addr . ' 2>&1', $output, $retval);
	echo "bakka\n";
	var_dump($output);
	var_dump($retval);
	print_r($output);

	//$obj = new COM ( 'WScript.Shell' );
	//$obj->Run('cmd /C wscript.exe ' . $file_addr, 0, false);
	//system('cscript.exe \"C:\xampp\htdocs\online_elink\processed_files\rbg_master16312720032604.xlsm.vbs\"');

	//exec( 'net user', $output, $retval );
	//exec( 'whoami', $output, $retval );
	
	//try {
	//$error_log = 'C:\xampp\htdocs\online_elink\error.log';
	//exec($file_addr.=" > $error_log 2>&1", $output, $retval);
	//echo "Returned with status $retval and output:\n";
	//print_r($output);
	//} catch (Exception $e) {
	//	echo 'Caught exception: ',  $e->getMessage(), "\n";
	//	print_r($e->getMessage());
	//}

	//system("cmd /c cscript ".$file_addr);  
	// system($file_addr);

	// $xlsObj = new COM("Excel.Application") or Die ("Did not connect");
	// $xlsObj->DisplayAlerts = false; 
	// $xlsObj->Workbooks->Open($file_addr);
	// $xlsObj->Run("GetData");
	// $excel->Quit;

	// $excel = new COM("Excel.Application") or die("Unable to instantiate excel");
	// $excel->Workbooks->Open($file_addr);
	// // run Excel silently, since we don’t want dialog boxes popping up in the background
	// $excel->DisplayAlerts = false;
	// $excel->Run("GetData");

	//3. call elink by command line
	// $output=null;
	// $retval=null;
	// $start_time = microtime(true);
	// print_r ($start_time);
	// print_r ($file_addr);
	// // exec( 'cmd /C wscript.exe '.$file_addr, $output, $retval );
	// system('cscript.exe "'.$file_addr.'"');
	// print_r ($output);
	// print_r ($retval);
	// $end_time = microtime(true);

	// $time_usage = ($end_time - $start_time);

//5. serve output value

} catch (Exception $e){
	header("HTTP/1.1 500 Internal Server Error");
	echo $e->getMessage();
	die();
}
?>