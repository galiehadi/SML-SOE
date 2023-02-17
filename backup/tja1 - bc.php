<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
header('Content-Type: application/json');

//get input data
$data_input = json_decode(file_get_contents('php://input'), true);


//1. copy file master
list($usec, $sec) = explode(" ", microtime());

$fileRep = 'tja1_'.((float)$usec + (float)$sec);
$fileRep = str_replace(".","",$fileRep);
$inputFileName = $fileRep.'.xlsm';
copy("master_tja1.xlsm",'processed_files\\'.$inputFileName);

//2. write value to file
//generate file
$vbs_content = 'Set objExcel_'.$fileRep.'    = CreateObject("Excel.Application")'."\r\n";
$vbs_content = $vbs_content. 'Set objWorkbook_'.$fileRep.' = objExcel_'.$fileRep.'.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'")'."\r\n";

// set value input start
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(256,5).Value = "'.$data_input['steam_mass_flow'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(509,5).Value = "'.$data_input['steam_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(429,5).Value = "'.$data_input['steam_temp'].'"'."\r\n"; 
// set value input end

$vbs_content = $vbs_content. 'objWorkbook_'.$fileRep.'.Save'."\r\n";
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Application.DisplayAlerts = False'."\r\n";
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Application.Run "\'C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'\'!GetData"'."\r\n";
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Application.DisplayAlerts = False'."\r\n";
$vbs_content = $vbs_content. 'objWorkbook_'.$fileRep.'.Save'."\r\n";
$vbs_content = $vbs_content. 'objWorkbook_'.$fileRep.'.Close False '."\r\n";
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Application.Quit'."\r\n";
$vbs_content = $vbs_content. 'Set objExcel_'.$fileRep.' = Nothing '."\r\n";
$vbs_content = $vbs_content. 'Set objWorkbook_'.$fileRep.' = Nothing'."\r\n";

// write to vbs file
$file_addr = 'C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'.vbs';
$vbs_file = fopen($file_addr, "wb") or die("Unable to open file!");
fwrite($vbs_file, $vbs_content);
fclose($vbs_file);

//3. call elink by command line
$output=null;
$retval=null;
exec( $file_addr, $output, $retval);

//4. get output value
/** Load $inputFileName to a Spreadsheet object **/
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('processed_files\\'.$inputFileName);

//defaine output start
$json_subdata1['boiler_eff']['value'] = $spreadsheet->getActiveSheet()->getCell('E1986')->getValue();
$json_subdata1['boiler_eff']['unit'] = $spreadsheet->getActiveSheet()->getCell('C1986')->getValue();

$json_subdata2['gross_power']['value'] = $spreadsheet->getActiveSheet()->getCell('E598')->getValue();
$json_subdata2['gross_power']['unit'] = $spreadsheet->getActiveSheet()->getCell('C598')->getValue();

$json_subdata3['net_power']['value'] = $spreadsheet->getActiveSheet()->getCell('E599')->getValue();
$json_subdata3['net_power']['unit'] = $spreadsheet->getActiveSheet()->getCell('C599')->getValue();

$json_subdata4['nphr']['value'] = $spreadsheet->getActiveSheet()->getCell('E606')->getValue();
$json_subdata4['nphr']['unit'] = $spreadsheet->getActiveSheet()->getCell('C606')->getValue();

$json_subdata5['gphr']['value'] = $spreadsheet->getActiveSheet()->getCell('E604')->getValue();
$json_subdata5['gphr']['unit'] = $spreadsheet->getActiveSheet()->getCell('C604')->getValue();
//defaine output end

$json_data = [$json_subdata1,$json_subdata2,$json_subdata3,$json_subdata4,$json_subdata5];

//5. serve output value
echo json_encode($json_data);

?>