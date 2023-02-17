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
//1. copy file master
list($usec, $sec) = explode(" ", microtime());

$fileRep = 'blk_master'.((float)$usec + (float)$sec);
$fileRep = str_replace(".","",$fileRep);
$inputFileName = $fileRep.'.xlsm';
copy("blk_master.xlsm",'processed_files\\'.$inputFileName);

//2. write value to file
//generate file
$vbs_content = 'Set objExcel_'.$fileRep.'    = CreateObject("Excel.Application")'."\r\n";
$vbs_content = $vbs_content. 'Set objWorkbook_'.$fileRep.' = objExcel_'.$fileRep.'.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'")'."\r\n";

// set value input start
//Worksheets("Input").Activate
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!C1").Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 

$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J4").Value = "'.($data_input['expected_power_output']*1000).'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J3").Value = "'.$data_input['generator_power_factor'].'"'."\r\n"; 
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(238,5).Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J5").Value = "'.($data_input['desired_coal_hhv']*4.1868).'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J6").Value = "'.$data_input['total_moisture'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J7").Value = "'.$data_input['ambient_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J8").Value = "'.$data_input['relative_humidity'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J9").Value = "'.$data_input['ambient_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J10").Value = "'.$data_input['excess_air'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J11").Value = "'.$data_input['water_cooling_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!J12").Value = "'.$data_input['cooling_temp_rise'].'"'."\r\n"; 


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
$start_time = microtime(true);
exec( $file_addr, $output, $retval );
// system('cscript.exe "'.$file_addr.'"');
$end_time = microtime(true);

$time_usage = ($end_time - $start_time);



//4. get output value
/** Load $inputFileName to a Spreadsheet object **/
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('processed_files\\'.$inputFileName);

// $elink_status = $spreadsheet->getSheetByName('output')->getCell('D6')->getCalculatedValue();
// if($elink_status = 'Failed'){
// 	header("HTTP/1.1 500 Internal Server Error");
// 	echo 'Thermodynamic calculation failed.';
// 	die();
// }

$spreadsheet->setActiveSheetIndex(5);

//defaine output start
$ambient_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('G3')->getCalculatedValue(),2);
$ambient_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('E3')->getCalculatedValue();
$relative_humidity_value = round($spreadsheet->getSheetByName('output')->getCell('G4')->getCalculatedValue(),2);
$relative_humidity_unit = $spreadsheet->getSheetByName('output')->getCell('E4')->getCalculatedValue();
$ambient_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('G5')->getCalculatedValue(),2);
$ambient_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('E5')->getCalculatedValue();
$gross_power_output_value = round($spreadsheet->getSheetByName('output')->getCell('G6')->getCalculatedValue(),2);
$gross_power_output_unit = $spreadsheet->getSheetByName('output')->getCell('E6')->getCalculatedValue();
$stack_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('G7')->getCalculatedValue(),2);
$stack_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('E7')->getCalculatedValue();
$flue_gas_inlet_ah_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G8')->getCalculatedValue(),2);
$flue_gas_inlet_ah_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E8')->getCalculatedValue();
$flue_gas_outlet_ah_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G9')->getCalculatedValue(),2);
$flue_gas_outlet_ah_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E9')->getCalculatedValue();
$pa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G10')->getCalculatedValue(),2);
$pa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E10')->getCalculatedValue();
$pa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G11')->getCalculatedValue(),2);
$pa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E11')->getCalculatedValue();
$sa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G12')->getCalculatedValue(),2);
$sa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E12')->getCalculatedValue();
$sa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G13')->getCalculatedValue(),2);
$sa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E13')->getCalculatedValue();
$main_steam_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('G14')->getCalculatedValue(),2);
$main_steam_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('E14')->getCalculatedValue();
$main_steam_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G15')->getCalculatedValue(),2);
$main_steam_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E15')->getCalculatedValue();
$main_steam_press_value = round($spreadsheet->getSheetByName('output')->getCell('G16')->getCalculatedValue(),2);
$main_steam_press_unit = $spreadsheet->getSheetByName('output')->getCell('E16')->getCalculatedValue();
$feedwater_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('G17')->getCalculatedValue(),2);
$feedwater_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('E17')->getCalculatedValue();
$feedwater_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G18')->getCalculatedValue(),2);
$feedwater_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E18')->getCalculatedValue();
$feedwater_press_value = round($spreadsheet->getSheetByName('output')->getCell('G19')->getCalculatedValue(),2);
$feedwater_press_unit = $spreadsheet->getSheetByName('output')->getCell('E19')->getCalculatedValue();
$condenser_press_value = round($spreadsheet->getSheetByName('output')->getCell('G20')->getCalculatedValue(),2);
$condenser_press_unit = $spreadsheet->getSheetByName('output')->getCell('E20')->getCalculatedValue();
$cooling_water_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G21')->getCalculatedValue(),2);
$cooling_water_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E21')->getCalculatedValue();
$cooling_water_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('G22')->getCalculatedValue(),2);
$cooling_water_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('E22')->getCalculatedValue();
$cooling_water_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('G23')->getCalculatedValue(),2);
$cooling_water_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('E23')->getCalculatedValue();
$coal_flow_value = round($spreadsheet->getSheetByName('output')->getCell('G24')->getCalculatedValue(),2);
$coal_flow_unit = $spreadsheet->getSheetByName('output')->getCell('E24')->getCalculatedValue();
$coal_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('G25')->getCalculatedValue(),2);
$coal_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('E25')->getCalculatedValue();
$turbine_1_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('G26')->getCalculatedValue(),2);
$turbine_1_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('E26')->getCalculatedValue();
$turbine_2_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('G27')->getCalculatedValue(),2);
$turbine_2_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('E27')->getCalculatedValue();
$turbine_3_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('G28')->getCalculatedValue(),2);
$turbine_3_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('E28')->getCalculatedValue();
$turbine_4_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('G29')->getCalculatedValue(),2);
$turbine_4_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('E29')->getCalculatedValue();
$turbine_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('G30')->getCalculatedValue(),2);
$turbine_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('E30')->getCalculatedValue();

$hph_dca_value = round($spreadsheet->getSheetByName('output')->getCell('G32')->getCalculatedValue(),2);
$hph_dca_unit = $spreadsheet->getSheetByName('output')->getCell('E32')->getCalculatedValue();
$hph_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('G33')->getCalculatedValue(),2);
$hph_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('E33')->getCalculatedValue();
$lph_dca_value = round($spreadsheet->getSheetByName('output')->getCell('G34')->getCalculatedValue(),2);
$lph_dca_unit = $spreadsheet->getSheetByName('output')->getCell('E34')->getCalculatedValue();
$lph_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('G35')->getCalculatedValue(),2);
$lph_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('E35')->getCalculatedValue();


$gross_output_value = round($spreadsheet->getSheetByName('output')->getCell('G38')->getCalculatedValue(),2);
$gross_output_unit = $spreadsheet->getSheetByName('output')->getCell('E38')->getCalculatedValue();
$gross_heat_rate_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('G39')->getCalculatedValue(),2);
$gross_heat_rate_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('E39')->getCalculatedValue();
$nett_output_value = round($spreadsheet->getSheetByName('output')->getCell('G40')->getCalculatedValue(),2);
$nett_output_unit = $spreadsheet->getSheetByName('output')->getCell('E40')->getCalculatedValue();
$nett_heat_rate_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('G41')->getCalculatedValue(),2);
$nett_heat_rate_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('E41')->getCalculatedValue();


$gross_heat_rate_lhv_value = round($spreadsheet->getSheetByName('output')->getCell('G44')->getCalculatedValue(),2);
$gross_heat_rate_lhv_unit = $spreadsheet->getSheetByName('output')->getCell('E44')->getCalculatedValue();
$hhv_value = round($spreadsheet->getSheetByName('output')->getCell('G45')->getCalculatedValue(),2);
$hhv_unit = $spreadsheet->getSheetByName('output')->getCell('E45')->getCalculatedValue();
$lhv_value = round($spreadsheet->getSheetByName('output')->getCell('G46')->getCalculatedValue(),2);
$lhv_unit = $spreadsheet->getSheetByName('output')->getCell('E46')->getCalculatedValue();


//interface
$json_subdata_int['gross_power_output']['value'] = round($spreadsheet->getSheetByName('output')->getCell('G6')->getCalculatedValue(),2);
$json_subdata_int['gross_power_output']['unit'] = $spreadsheet->getSheetByName('output')->getCell('E63')->getCalculatedValue();

// print_r('gross heat rate ',round($spreadsheet->getSheetByName('output')->getCell('G39')->getCalculatedValue(),2));

$json_subdata_int['gross_heatrate']['value'] = round($spreadsheet->getSheetByName('output')->getCell('G42')->getCalculatedValue(),2);
$json_subdata_int['gross_heatrate']['unit'] = $spreadsheet->getSheetByName('output')->getCell('E39')->getCalculatedValue();
$json_subdata_int['net_output']['value'] = round($spreadsheet->getSheetByName('output')->getCell('G40')->getCalculatedValue(),2);
$json_subdata_int['net_output']['unit'] = $spreadsheet->getSheetByName('output')->getCell('E40')->getCalculatedValue();
$json_subdata_int['net_heatrate']['value'] = round($spreadsheet->getSheetByName('output')->getCell('G41')->getCalculatedValue(),2);
$json_subdata_int['net_heatrate']['unit'] = $spreadsheet->getSheetByName('output')->getCell('E41')->getCalculatedValue();
// $json_subdata_int['gross_heatrate']['value'] = round($spreadsheet->getSheetByName('output')->getCell('F69')->getCalculatedValue(),2);
// $json_subdata_int['gross_heatrate']['unit'] = $spreadsheet->getSheetByName('output')->getCell('D69')->getCalculatedValue();
$json_subdata_int['time_usage']['value'] = round($time_usage,3); 
$json_subdata_int['time_usage']['unit'] = 'seconds';

// $json_subdata_int['hhv']['value'] = round($spreadsheet->getSheetByName('output')->getCell('F70')->getCalculatedValue(),2);
// $json_subdata_int['gross_power_output']['unit'] = $spreadsheet->getSheetByName('output')->getCell('D70')->getCalculatedValue();
// $json_subdata_int['lhv']['value'] = round($spreadsheet->getSheetByName('output')->getCell('F71')->getCalculatedValue(),2);
// $json_subdata_int['gross_power_output']['unit'] = $spreadsheet->getSheetByName('output')->getCell('D71')->getCalculatedValue();

$json_data['interface'] = $json_subdata_int;

$json_subdata_1['name'] = 'Ambient Temperature';
$json_subdata_1['unit'] = $ambient_temperature_unit;
$json_subdata_1['value'] = $ambient_temperature_value;
$json_subdata_2['name'] = 'Relative Humidity';
$json_subdata_2['unit'] = $relative_humidity_unit;
$json_subdata_2['value'] = $relative_humidity_value;
$json_subdata_3['name'] = 'Ambient Pressure';
$json_subdata_3['unit'] = $ambient_pressure_unit;
$json_subdata_3['value'] = $ambient_pressure_value;
$json_subdata_4['name'] = 'Gross Power Output';
$json_subdata_4['unit'] = $gross_power_output_unit;
$json_subdata_4['value'] = $gross_power_output_value;
$json_subdata_5['name'] = 'Stack Temperature';
$json_subdata_5['unit'] = $stack_temperature_unit;
$json_subdata_5['value'] = $stack_temperature_value;
$json_subdata_6['name'] = 'Flue Gas Inlet AH Temp.';
$json_subdata_6['unit'] = $flue_gas_inlet_ah_temp_unit;
$json_subdata_6['value'] = $flue_gas_inlet_ah_temp_value;
$json_subdata_7['name'] = 'Flue Gas Outlet AH Temp.';
$json_subdata_7['unit'] = $flue_gas_outlet_ah_temp_unit;
$json_subdata_7['value'] = $flue_gas_outlet_ah_temp_value;
$json_subdata_8['name'] = 'PA Inlet Temp.';
$json_subdata_8['unit'] = $pa_inlet_temp_unit;
$json_subdata_8['value'] = $pa_inlet_temp_value;
$json_subdata_9['name'] = 'PA Outlet Temp.';
$json_subdata_9['unit'] = $pa_outlet_temp_unit;
$json_subdata_9['value'] = $pa_outlet_temp_value;
$json_subdata_10['name'] = 'SA Inlet Temp.';
$json_subdata_10['unit'] = $sa_inlet_temp_unit;
$json_subdata_10['value'] = $sa_inlet_temp_value;
$json_subdata_11['name'] = 'SA Outlet Temp.';
$json_subdata_11['unit'] = $sa_outlet_temp_unit;
$json_subdata_11['value'] = $sa_outlet_temp_value;
$json_subdata_12['name'] = 'Main Steam Mass Flow';
$json_subdata_12['unit'] = $main_steam_mass_flow_unit;
$json_subdata_12['value'] = $main_steam_mass_flow_value;
$json_subdata_13['name'] = 'Main Steam Temp';
$json_subdata_13['unit'] = $main_steam_temp_unit;
$json_subdata_13['value'] = $main_steam_temp_value;
$json_subdata_14['name'] = 'Main Steam Press.';
$json_subdata_14['unit'] = $main_steam_press_unit;
$json_subdata_14['value'] = $main_steam_press_value;
$json_subdata_15['name'] = 'Feedwater Mass Flow';
$json_subdata_15['unit'] = $feedwater_mass_flow_unit;
$json_subdata_15['value'] = $feedwater_mass_flow_value;
$json_subdata_16['name'] = 'Feedwater Temp';
$json_subdata_16['unit'] = $feedwater_temp_unit;
$json_subdata_16['value'] = $feedwater_temp_value;
$json_subdata_17['name'] = 'Feedwater Press';
$json_subdata_17['unit'] = $feedwater_press_unit;
$json_subdata_17['value'] = $feedwater_press_value;
$json_subdata_18['name'] = 'Condenser Press';
$json_subdata_18['unit'] = $condenser_press_unit;
$json_subdata_18['value'] = $condenser_press_value;
$json_subdata_19['name'] = 'Cooling Water Inlet Temp';
$json_subdata_19['unit'] = $cooling_water_inlet_temp_unit;
$json_subdata_19['value'] = $cooling_water_inlet_temp_value;
$json_subdata_20['name'] = 'Cooling Water Mass Flow';
$json_subdata_20['unit'] = $cooling_water_mass_flow_unit;
$json_subdata_20['value'] = $cooling_water_mass_flow_value;
$json_subdata_21['name'] = 'Cooling Water Outlet Temp.';
$json_subdata_21['unit'] = $cooling_water_outlet_temp_unit;
$json_subdata_21['value'] = $cooling_water_outlet_temp_value;
$json_subdata_22['name'] = 'Coal Flow';
$json_subdata_22['unit'] = $coal_flow_unit;
$json_subdata_22['value'] = $coal_flow_value;
$json_subdata_23['name'] = 'Coal HHV';
$json_subdata_23['unit'] = $coal_hhv_unit;
$json_subdata_23['value'] = $coal_hhv_value;
$json_subdata_24['name'] = 'Turbine Efficiency';
$json_subdata_24['unit'] = $turbine_efficiency_unit;
$json_subdata_24['value'] = $turbine_efficiency_value;

$json_subdata_26['name'] = 'HPH DCA';
$json_subdata_26['unit'] = $hph_dca_unit;
$json_subdata_26['value'] = $hph_dca_value;
$json_subdata_27['name'] = 'HPH TTD';
$json_subdata_27['unit'] = $hph_ttd_unit;
$json_subdata_27['value'] = $hph_ttd_value;
$json_subdata_28['name'] = 'LPH DCA';
$json_subdata_28['unit'] = $lph_dca_unit;
$json_subdata_28['value'] = $lph_dca_value;
$json_subdata_29['name'] = 'LPH TTD';
$json_subdata_29['unit'] = $lph_ttd_unit;
$json_subdata_29['value'] = $lph_ttd_value;


$json_subdata_30['name'] = 'Gross Output';
$json_subdata_30['unit'] = $gross_output_unit;
$json_subdata_30['value'] = $gross_output_value;
$json_subdata_31['name'] = 'Gross Heat Rate HHV';
$json_subdata_31['unit'] = $gross_heat_rate_hhv_unit;
$json_subdata_31['value'] = $gross_heat_rate_hhv_value;
$json_subdata_32['name'] = 'Nett Output';
$json_subdata_32['unit'] = $nett_output_unit;
$json_subdata_32['value'] = $nett_output_value;
$json_subdata_33['name'] = 'Nett Heat Rate HHV';
$json_subdata_33['unit'] = $nett_heat_rate_hhv_unit;
$json_subdata_33['value'] = $nett_heat_rate_hhv_value;


$json_subdata_34['name'] = 'Gross Heat Rate LHV';
$json_subdata_34['unit'] = $gross_heat_rate_lhv_unit;
$json_subdata_34['value'] = $gross_heat_rate_lhv_value;
$json_subdata_35['name'] = 'HHV';
$json_subdata_35['unit'] = $hhv_unit;
$json_subdata_35['value'] = $hhv_value;
$json_subdata_36['name'] = 'LHV';
$json_subdata_36['unit'] = $lhv_unit;
$json_subdata_36['value'] = $lhv_value;




$json_data['export'] = [
$json_subdata_1,
$json_subdata_2,
$json_subdata_3,
$json_subdata_4,
$json_subdata_5,
$json_subdata_6,
$json_subdata_7,
$json_subdata_8,
$json_subdata_9,
$json_subdata_10,
$json_subdata_11,
$json_subdata_12,
$json_subdata_13,
$json_subdata_14,
$json_subdata_15,
$json_subdata_16,
$json_subdata_17,
$json_subdata_18,
$json_subdata_19,
$json_subdata_20,
$json_subdata_21,
$json_subdata_22,
$json_subdata_23,
$json_subdata_24,
$json_subdata_26,
$json_subdata_27,
$json_subdata_28,
$json_subdata_29,
$json_subdata_30,
$json_subdata_31,
$json_subdata_32,
$json_subdata_33,
$json_subdata_34,
$json_subdata_35,
$json_subdata_36
];

$json_draw_1['id'] = 'D1';
$json_draw_1['value'] = $ambient_temperature_value. ' '. $ambient_temperature_unit;
$json_draw_2['id'] = 'D2';
$json_draw_2['value'] = $relative_humidity_value. ' '. $relative_humidity_unit;
$json_draw_3['id'] = 'D3';
$json_draw_3['value'] = $ambient_pressure_value. ' '. $ambient_pressure_unit;
$json_draw_4['id'] = 'D4';
$json_draw_4['value'] = $gross_power_output_value. ' '. $gross_power_output_unit;
$json_draw_5['id'] = 'D5';
$json_draw_5['value'] = $stack_temperature_value. ' '. $stack_temperature_unit;
$json_draw_6['id'] = 'D6';
$json_draw_6['value'] = $flue_gas_inlet_ah_temp_value. ' '. $flue_gas_inlet_ah_temp_unit;
$json_draw_7['id'] = 'D7';
$json_draw_7['value'] = $flue_gas_outlet_ah_temp_value. ' '. $flue_gas_outlet_ah_temp_unit;
$json_draw_8['id'] = 'D8';
$json_draw_8['value'] = $pa_inlet_temp_value. ' '. $pa_inlet_temp_unit;
$json_draw_9['id'] = 'D9';
$json_draw_9['value'] = $pa_outlet_temp_value. ' '. $pa_outlet_temp_unit;
$json_draw_10['id'] = 'D10';
$json_draw_10['value'] = $sa_inlet_temp_value. ' '. $sa_inlet_temp_unit;
$json_draw_11['id'] = 'D11';
$json_draw_11['value'] = $sa_outlet_temp_value. ' '. $sa_outlet_temp_unit;
$json_draw_12['id'] = 'D12';
$json_draw_12['value'] = $main_steam_mass_flow_value. ' '. $main_steam_mass_flow_unit;
$json_draw_13['id'] = 'D13';
$json_draw_13['value'] = $main_steam_temp_value. ' '. $main_steam_temp_unit;
$json_draw_14['id'] = 'D14';
$json_draw_14['value'] = $main_steam_press_value. ' '. $main_steam_press_unit;
$json_draw_15['id'] = 'D15';
$json_draw_15['value'] = $feedwater_mass_flow_value. ' '. $feedwater_mass_flow_unit;
$json_draw_16['id'] = 'D16';
$json_draw_16['value'] = $feedwater_temp_value. ' '. $feedwater_temp_unit;
$json_draw_17['id'] = 'D17';
$json_draw_17['value'] = $feedwater_press_value. ' '. $feedwater_press_unit;
$json_draw_18['id'] = 'D18';
$json_draw_18['value'] = $condenser_press_value. ' '. $condenser_press_unit;
$json_draw_19['id'] = 'D19';
$json_draw_19['value'] = $cooling_water_inlet_temp_value. ' '. $cooling_water_inlet_temp_unit;
$json_draw_20['id'] = 'D20';
$json_draw_20['value'] = $cooling_water_mass_flow_value. ' '. $cooling_water_mass_flow_unit;
$json_draw_21['id'] = 'D21';
$json_draw_21['value'] = $cooling_water_outlet_temp_value. ' '. $cooling_water_outlet_temp_unit;
$json_draw_22['id'] = 'D22';
$json_draw_22['value'] = $coal_flow_value. ' '. $coal_flow_unit;
$json_draw_23['id'] = 'D23';
$json_draw_23['value'] = $coal_hhv_value. ' '. $coal_hhv_unit;
$json_draw_24['id'] = 'D24';
$json_draw_24['value'] = $turbine_efficiency_value. ' '. $turbine_efficiency_unit;
$json_draw_26['id'] = 'D26';
$json_draw_26['value'] = $hph_dca_value. ' '. $hph_dca_unit;
$json_draw_27['id'] = 'D27';
$json_draw_27['value'] = $hph_ttd_value. ' '. $hph_ttd_unit;
$json_draw_28['id'] = 'D28';
$json_draw_28['value'] = $lph_dca_value. ' '. $lph_dca_unit;
$json_draw_29['id'] = 'D29';
$json_draw_29['value'] = $lph_ttd_value. ' '. $lph_ttd_unit;
$json_draw_30['id'] = 'D30';
$json_draw_30['value'] = $gross_output_value. ' '. $gross_output_unit;
$json_draw_31['id'] = 'D31';
$json_draw_31['value'] = $gross_heat_rate_hhv_value. ' '. $gross_heat_rate_hhv_unit;
$json_draw_32['id'] = 'D32';
$json_draw_32['value'] = $nett_output_value. ' '. $nett_output_unit;
$json_draw_33['id'] = 'D33';
$json_draw_33['value'] = $nett_heat_rate_hhv_value. ' '. $nett_heat_rate_hhv_unit;
$json_draw_34['id'] = 'D34';
$json_draw_34['value'] = $gross_heat_rate_lhv_value. ' '. $gross_heat_rate_lhv_unit;
$json_draw_35['id'] = 'D35';
$json_draw_35['value'] = $hhv_value. ' '. $hhv_unit;
$json_draw_36['id'] = 'D36';
$json_draw_36['value'] = $lhv_value. ' '. $lhv_unit;


$json_data['drawing'] = [

$json_draw_1,
$json_draw_2,
$json_draw_3,
$json_draw_4,
$json_draw_5,
$json_draw_6,
$json_draw_7,
$json_draw_8,
$json_draw_9,
$json_draw_10,
$json_draw_11,
$json_draw_12,
$json_draw_13,
$json_draw_14,
$json_draw_15,
$json_draw_16,
$json_draw_17,
$json_draw_18,
$json_draw_19,
$json_draw_20,
$json_draw_21,
$json_draw_22,
$json_draw_23,
$json_draw_24,
$json_draw_26,
$json_draw_27,
$json_draw_28,
$json_draw_29,
$json_draw_30,
$json_draw_31,
$json_draw_32,
$json_draw_33,
$json_draw_34,
$json_draw_35,
$json_draw_36
];

//5. serve output value output value will skip after error.
echo json_encode($json_data);
} catch (Exception $e){
	header("HTTP/1.1 500 Internal Server Error");
	echo $e->getMessage();
	die();
}
?>