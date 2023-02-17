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

$fileRep = 'ktt_master'.((float)$usec + (float)$sec);
$fileRep = str_replace(".","",$fileRep);
$inputFileName = $fileRep.'.xlsm';
copy("ktt_master.xlsm",'processed_files\\'.$inputFileName);

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

// $elink_status = $spreadsheet->getSheetByName('output')->getCell('E6')->getCalculatedValue();
// if($elink_status = 'Failed'){
// 	header("HTTP/1.1 500 Internal Server Error");
// 	echo 'Thermodynamic calculation failed.';
// 	die();
// }

$spreadsheet->setActiveSheetIndex(5);

//defaine output start
$ambient_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('F3')->getCalculatedValue(),2);
$ambient_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('D3')->getCalculatedValue();
$relative_humidity_value = round($spreadsheet->getSheetByName('output')->getCell('F4')->getCalculatedValue(),2);
$relative_humidity_unit = $spreadsheet->getSheetByName('output')->getCell('D4')->getCalculatedValue();
$ambient_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('F5')->getCalculatedValue(),2);
$ambient_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('D5')->getCalculatedValue();
$gross_power_output_value = round($spreadsheet->getSheetByName('output')->getCell('F6')->getCalculatedValue(),2);
$gross_power_output_unit = $spreadsheet->getSheetByName('output')->getCell('D6')->getCalculatedValue();
$stack_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('F7')->getCalculatedValue(),2);
$stack_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('D7')->getCalculatedValue();
$flue_gas_inlet_ah_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F8')->getCalculatedValue(),2);
$flue_gas_inlet_ah_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D8')->getCalculatedValue();
$flue_gas_outlet_ah_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F9')->getCalculatedValue(),2);
$flue_gas_outlet_ah_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D9')->getCalculatedValue();
$pa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F10')->getCalculatedValue(),2);
$pa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D10')->getCalculatedValue();
$pa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F11')->getCalculatedValue(),2);
$pa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D11')->getCalculatedValue();
$sa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F12')->getCalculatedValue(),2);
$sa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D12')->getCalculatedValue();
$sa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F13')->getCalculatedValue(),2);
$sa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D13')->getCalculatedValue();
$main_steam_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('F14')->getCalculatedValue(),2);
$main_steam_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('D14')->getCalculatedValue();
$main_steam_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F15')->getCalculatedValue(),2);
$main_steam_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D15')->getCalculatedValue();
$main_steam_press_value = round($spreadsheet->getSheetByName('output')->getCell('F16')->getCalculatedValue(),2);
$main_steam_press_unit = $spreadsheet->getSheetByName('output')->getCell('D16')->getCalculatedValue();
$feedwater_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('F17')->getCalculatedValue(),2);
$feedwater_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('D17')->getCalculatedValue();
$feedwater_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F18')->getCalculatedValue(),2);
$feedwater_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D18')->getCalculatedValue();
$feedwater_press_value = round($spreadsheet->getSheetByName('output')->getCell('F19')->getCalculatedValue(),2);
$feedwater_press_unit = $spreadsheet->getSheetByName('output')->getCell('D19')->getCalculatedValue();
$condenser_press_value = round($spreadsheet->getSheetByName('output')->getCell('F20')->getCalculatedValue(),2);
$condenser_press_unit = $spreadsheet->getSheetByName('output')->getCell('D20')->getCalculatedValue();
$cooling_water_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F21')->getCalculatedValue(),2);
$cooling_water_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D21')->getCalculatedValue();
$cooling_water_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('F22')->getCalculatedValue(),2);
$cooling_water_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('D22')->getCalculatedValue();
$cooling_water_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F23')->getCalculatedValue(),2);
$cooling_water_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D23')->getCalculatedValue();
$coal_flow_value = round($spreadsheet->getSheetByName('output')->getCell('F24')->getCalculatedValue(),2);
$coal_flow_unit = $spreadsheet->getSheetByName('output')->getCell('D24')->getCalculatedValue();
$coal_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('F25')->getCalculatedValue(),2);
$coal_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('D25')->getCalculatedValue();
$lp_steam_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('F26')->getCalculatedValue(),2);
$lp_steam_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('D26')->getCalculatedValue();
$lp_steam_temp_value = round($spreadsheet->getSheetByName('output')->getCell('F27')->getCalculatedValue(),2);
$lp_steam_temp_unit = $spreadsheet->getSheetByName('output')->getCell('D27')->getCalculatedValue();
$lp_steam_press_value = round($spreadsheet->getSheetByName('output')->getCell('F28')->getCalculatedValue(),2);
$lp_steam_press_unit = $spreadsheet->getSheetByName('output')->getCell('D28')->getCalculatedValue();
$st_group_1_hpt_1_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F29')->getCalculatedValue(),2);
$st_group_1_hpt_1_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D29')->getCalculatedValue();
$st_group_2_hpt_2_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F30')->getCalculatedValue(),2);
$st_group_2_hpt_2_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D30')->getCalculatedValue();
$st_group_3_hpt_3_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F31')->getCalculatedValue(),2);
$st_group_3_hpt_3_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D31')->getCalculatedValue();
$st_group_4_hpt_4_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F32')->getCalculatedValue(),2);
$st_group_4_hpt_4_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D32')->getCalculatedValue();
$st_group_5_hpt_5_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F33')->getCalculatedValue(),2);
$st_group_5_hpt_5_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D33')->getCalculatedValue();
$st_group_6_lpt_1_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F34')->getCalculatedValue(),2);
$st_group_6_lpt_1_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D34')->getCalculatedValue();
$st_group_7_lpt_2_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F35')->getCalculatedValue(),2);
$st_group_7_lpt_2_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D35')->getCalculatedValue();
$st_group_8_lpt_3_dry_step_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F36')->getCalculatedValue(),2);
$st_group_8_lpt_3_dry_step_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D36')->getCalculatedValue();
$hp_turbine_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F37')->getCalculatedValue(),2);
$hp_turbine_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D37')->getCalculatedValue();
$lp_turbine_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('F38')->getCalculatedValue(),2);
$lp_turbine_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('D38')->getCalculatedValue();

$hph_1_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F40')->getCalculatedValue(),2);
$hph_1_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D40')->getCalculatedValue();
$hph_1_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F41')->getCalculatedValue(),2);
$hph_1_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D41')->getCalculatedValue();
$hph_2_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F42')->getCalculatedValue(),2);
$hph_2_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D42')->getCalculatedValue();
$hph_2_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F43')->getCalculatedValue(),2);
$hph_2_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D43')->getCalculatedValue();
$lph_4_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F44')->getCalculatedValue(),2);
$lph_4_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D44')->getCalculatedValue();
$lph_4_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F45')->getCalculatedValue(),2);
$lph_4_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D45')->getCalculatedValue();
$lph_5_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F46')->getCalculatedValue(),2);
$lph_5_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D46')->getCalculatedValue();
$lph_5_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F47')->getCalculatedValue(),2);
$lph_5_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D47')->getCalculatedValue();
$lph_6_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F48')->getCalculatedValue(),2);
$lph_6_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D48')->getCalculatedValue();
$lph_6_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F49')->getCalculatedValue(),2);
$lph_6_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D49')->getCalculatedValue();
$lph_7_dca_value = round($spreadsheet->getSheetByName('output')->getCell('F50')->getCalculatedValue(),2);
$lph_7_dca_unit = $spreadsheet->getSheetByName('output')->getCell('D50')->getCalculatedValue();
$lph_7_ttd_value = round($spreadsheet->getSheetByName('output')->getCell('F51')->getCalculatedValue(),2);
$lph_7_ttd_unit = $spreadsheet->getSheetByName('output')->getCell('D51')->getCalculatedValue();

$gross_power_output_value = round($spreadsheet->getSheetByName('output')->getCell('F54')->getCalculatedValue(),2);
$gross_power_output_unit = $spreadsheet->getSheetByName('output')->getCell('D54')->getCalculatedValue();
$gross_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('F55')->getCalculatedValue(),2);
$gross_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('D55')->getCalculatedValue();
$net_output_value = round($spreadsheet->getSheetByName('output')->getCell('F56')->getCalculatedValue(),2);
$net_output_unit = $spreadsheet->getSheetByName('output')->getCell('D56')->getCalculatedValue();
$net_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('F57')->getCalculatedValue(),2);
$net_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('D57')->getCalculatedValue();



$json_subdata_int['gross_power_output']['value'] = $gross_power_output_value;
$json_subdata_int['gross_power_output']['unit'] = $gross_power_output_unit; 
$json_subdata_int['gross_heatrate']['value'] = $gross_heatrate_value;
$json_subdata_int['gross_heatrate']['unit'] = $gross_heatrate_unit;
$json_subdata_int['net_output']['value'] = $net_output_value;
$json_subdata_int['net_output']['unit'] = $net_output_unit;
$json_subdata_int['net_heatrate']['value'] = $net_heatrate_value;
$json_subdata_int['net_heatrate']['unit'] = $net_heatrate_unit;
$json_subdata_int['time_usage']['value'] = round($time_usage,3); 
$json_subdata_int['time_usage']['unit'] = 'seconds';

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
$json_subdata_13['name'] = 'Main Steam Temp.';
$json_subdata_13['unit'] = $main_steam_temp_unit;
$json_subdata_13['value'] = $main_steam_temp_value;
$json_subdata_14['name'] = 'Main Steam Press.';
$json_subdata_14['unit'] = $main_steam_press_unit;
$json_subdata_14['value'] = $main_steam_press_value;
$json_subdata_15['name'] = 'Feedwater Mass Flow';
$json_subdata_15['unit'] = $feedwater_mass_flow_unit;
$json_subdata_15['value'] = $feedwater_mass_flow_value;
$json_subdata_16['name'] = 'Feedwater Temp.';
$json_subdata_16['unit'] = $feedwater_temp_unit;
$json_subdata_16['value'] = $feedwater_temp_value;
$json_subdata_17['name'] = 'Feedwater Press.';
$json_subdata_17['unit'] = $feedwater_press_unit;
$json_subdata_17['value'] = $feedwater_press_value;
$json_subdata_18['name'] = 'Condenser Press.';
$json_subdata_18['unit'] = $condenser_press_unit;
$json_subdata_18['value'] = $condenser_press_value;
$json_subdata_19['name'] = 'Cooling Water Inlet Temp.';
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
$json_subdata_24['name'] = 'LP Steam Mass Flow';
$json_subdata_24['unit'] = $lp_steam_mass_flow_unit;
$json_subdata_24['value'] = $lp_steam_mass_flow_value;
$json_subdata_25['name'] = 'LP Steam Temp.';
$json_subdata_25['unit'] = $lp_steam_temp_unit;
$json_subdata_25['value'] = $lp_steam_temp_value;
$json_subdata_26['name'] = 'LP Steam Press.';
$json_subdata_26['unit'] = $lp_steam_press_unit;
$json_subdata_26['value'] = $lp_steam_press_value;
$json_subdata_27['name'] = 'ST Group [1] - HPT 1 : Dry step efficiency';
$json_subdata_27['unit'] = $st_group_1_hpt_1_dry_step_efficiency_unit;
$json_subdata_27['value'] = $st_group_1_hpt_1_dry_step_efficiency_value;
$json_subdata_28['name'] = 'ST Group [2] - HPT 2 : Dry step efficiency';
$json_subdata_28['unit'] = $st_group_2_hpt_2_dry_step_efficiency_unit;
$json_subdata_28['value'] = $st_group_2_hpt_2_dry_step_efficiency_value;
$json_subdata_29['name'] = 'ST Group [3] - HPT 3 : Dry step efficiency';
$json_subdata_29['unit'] = $st_group_3_hpt_3_dry_step_efficiency_unit;
$json_subdata_29['value'] = $st_group_3_hpt_3_dry_step_efficiency_value;
$json_subdata_30['name'] = 'ST Group [4] - HPT 4 : Dry step efficiency';
$json_subdata_30['unit'] = $st_group_4_hpt_4_dry_step_efficiency_unit;
$json_subdata_30['value'] = $st_group_4_hpt_4_dry_step_efficiency_value;
$json_subdata_31['name'] = 'ST Group [5] - HPT 5 : Dry step efficiency';
$json_subdata_31['unit'] = $st_group_5_hpt_5_dry_step_efficiency_unit;
$json_subdata_31['value'] = $st_group_5_hpt_5_dry_step_efficiency_value;
$json_subdata_32['name'] = 'ST Group [6] - LPT 1 : Dry step efficiency';
$json_subdata_32['unit'] = $st_group_6_lpt_1_dry_step_efficiency_unit;
$json_subdata_32['value'] = $st_group_6_lpt_1_dry_step_efficiency_value;
$json_subdata_33['name'] = 'ST Group [7] - LPT 2 : Dry step efficiency';
$json_subdata_33['unit'] = $st_group_7_lpt_2_dry_step_efficiency_unit;
$json_subdata_33['value'] = $st_group_7_lpt_2_dry_step_efficiency_value;
$json_subdata_34['name'] = 'ST Group [8] - LPT 3 : Dry step efficiency';
$json_subdata_34['unit'] = $st_group_8_lpt_3_dry_step_efficiency_unit;
$json_subdata_34['value'] = $st_group_8_lpt_3_dry_step_efficiency_value;
$json_subdata_35['name'] = 'HP Turbine Efficiency';
$json_subdata_35['unit'] = $hp_turbine_efficiency_unit;
$json_subdata_35['value'] = $hp_turbine_efficiency_value;
$json_subdata_36['name'] = 'LP Turbine Efficiency';
$json_subdata_36['unit'] = $lp_turbine_efficiency_unit;
$json_subdata_36['value'] = $lp_turbine_efficiency_value;

$json_subdata_38['name'] = 'HPH-1 DCA';
$json_subdata_38['unit'] = $hph_1_dca_unit;
$json_subdata_38['value'] = $hph_1_dca_value;
$json_subdata_39['name'] = 'HPH-1 TTD';
$json_subdata_39['unit'] = $hph_1_ttd_unit;
$json_subdata_39['value'] = $hph_1_ttd_value;
$json_subdata_40['name'] = 'HPH-2 DCA';
$json_subdata_40['unit'] = $hph_2_dca_unit;
$json_subdata_40['value'] = $hph_2_dca_value;
$json_subdata_41['name'] = 'HPH-2 TTD';
$json_subdata_41['unit'] = $hph_2_ttd_unit;
$json_subdata_41['value'] = $hph_2_ttd_value;
$json_subdata_42['name'] = 'LPH-4DCA';
$json_subdata_42['unit'] = $lph_4_dca_unit;
$json_subdata_42['value'] = $lph_4_dca_value;
$json_subdata_43['name'] = 'LPH-4 TTD';
$json_subdata_43['unit'] = $lph_4_ttd_unit;
$json_subdata_43['value'] = $lph_4_ttd_value;
$json_subdata_44['name'] = 'LPH-5 DCA';
$json_subdata_44['unit'] = $lph_5_dca_unit;
$json_subdata_44['value'] = $lph_5_dca_value;
$json_subdata_45['name'] = 'LPH-5 TTD';
$json_subdata_45['unit'] = $lph_5_ttd_unit;
$json_subdata_45['value'] = $lph_5_ttd_value;
$json_subdata_46['name'] = 'LPH-6 DCA';
$json_subdata_46['unit'] = $lph_6_dca_unit;
$json_subdata_46['value'] = $lph_6_dca_value;
$json_subdata_47['name'] = 'LPH-6 TTD';
$json_subdata_47['unit'] = $lph_6_ttd_unit;
$json_subdata_47['value'] = $lph_6_ttd_value;
$json_subdata_48['name'] = 'LPH-7 DCA';
$json_subdata_48['unit'] = $lph_7_dca_unit;
$json_subdata_48['value'] = $lph_7_dca_value;
$json_subdata_49['name'] = 'LPH-7 TTD';
$json_subdata_49['unit'] = $lph_7_ttd_unit;
$json_subdata_49['value'] = $lph_7_ttd_value;


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
$json_subdata_25,
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
$json_subdata_36,

$json_subdata_38,
$json_subdata_39,
$json_subdata_40,
$json_subdata_41,
$json_subdata_42,
$json_subdata_43,
$json_subdata_44,
$json_subdata_45,
$json_subdata_46,
$json_subdata_47,
$json_subdata_48,
$json_subdata_49
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
$json_draw_24['value'] = $lp_steam_mass_flow_value. ' '. $lp_steam_mass_flow_unit; 
$json_draw_25['id'] = 'D25';
$json_draw_25['value'] = $lp_steam_temp_value. ' '. $lp_steam_temp_unit; 
$json_draw_26['id'] = 'D26';
$json_draw_26['value'] = $lp_steam_press_value. ' '. $lp_steam_press_unit; 

$json_draw_35['id'] = 'D27';
$json_draw_35['value'] = $hp_turbine_efficiency_value. ' '. $hp_turbine_efficiency_unit; 
$json_draw_36['id'] = 'D28';
$json_draw_36['value'] = $lp_turbine_efficiency_value. ' '. $lp_turbine_efficiency_unit; 

$json_draw_38['id'] = 'D30';
$json_draw_38['value'] = $hph_1_dca_value. ' '. $hph_1_dca_unit; 
$json_draw_39['id'] = 'D31';
$json_draw_39['value'] = $hph_1_ttd_value. ' '. $hph_1_ttd_unit; 
$json_draw_40['id'] = 'D32';
$json_draw_40['value'] = $hph_2_dca_value. ' '. $hph_2_dca_unit; 
$json_draw_41['id'] = 'D33';
$json_draw_41['value'] = $hph_2_ttd_value. ' '. $hph_2_ttd_unit; 
$json_draw_42['id'] = 'D34';
$json_draw_42['value'] = $lph_4_dca_value. ' '. $lph_4_dca_unit; 
$json_draw_43['id'] = 'D35';
$json_draw_43['value'] = $lph_4_ttd_value. ' '. $lph_4_ttd_unit; 
$json_draw_44['id'] = 'D36';
$json_draw_44['value'] = $lph_5_dca_value. ' '. $lph_5_dca_unit; 
$json_draw_45['id'] = 'D37';
$json_draw_45['value'] = $lph_5_ttd_value. ' '. $lph_5_ttd_unit; 
$json_draw_46['id'] = 'D38';
$json_draw_46['value'] = $lph_6_dca_value. ' '. $lph_6_dca_unit; 
$json_draw_47['id'] = 'D39';
$json_draw_47['value'] = $lph_6_ttd_value. ' '. $lph_6_ttd_unit; 
$json_draw_48['id'] = 'D40';
$json_draw_48['value'] = $lph_7_dca_value. ' '. $lph_7_dca_unit; 
$json_draw_49['id'] = 'D41';
$json_draw_49['value'] = $lph_7_ttd_value. ' '. $lph_7_ttd_unit; 


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
	$json_draw_25,
	$json_draw_26,
	$json_draw_35,
	$json_draw_36,
	$json_draw_38,
	$json_draw_39,
	$json_draw_40,
	$json_draw_41,
	$json_draw_42,
	$json_draw_43,
	$json_draw_44,
	$json_draw_45,
	$json_draw_46,
	$json_draw_47,
	$json_draw_48,
	$json_draw_49
];

//5. serve output value output value will skip after error.
echo json_encode($json_data);
} catch (Exception $e){
	header("HTTP/1.1 500 Internal Server Error");
	echo $e->getMessage();
	die();
}
?>