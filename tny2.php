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

$fileRep = 'tny_master'.((float)$usec + (float)$sec);
$fileRep = str_replace(".","",$fileRep);
$inputFileName = $fileRep.'.xlsm';
copy("tny_master.xlsm",'processed_files\\'.$inputFileName);

//2. write value to file
//generate file
$vbs_content = 'Set objExcel_'.$fileRep.'    = CreateObject("Excel.Application")'."\r\n";
$vbs_content = $vbs_content. 'Set objWorkbook_'.$fileRep.' = objExcel_'.$fileRep.'.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'")'."\r\n";

// set value input start
//yang lama
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(488,5).Value = "'.$data_input['generator_power_factor'].'"'."\r\n"; 
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(238,5).Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(238,5).Value = "'.$data_input['desired_coal_hhv'].'"'."\r\n"; 
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(230,5).Value = "'.$data_input['tota// l_moisture'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(17,5).Value = "'.$data_input['ambient_temp'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(18,5).Value = "'.$data_input['relative_humidity'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(230,5).Value = "'.$data_input['total_moisture'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(20,5).Value = "'.$data_input['ambient_press'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(258,5).Value = "'.$data_input['excess_air'].'"'."\r\n";
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(429,5).Value = "538"'."\r\n"; 

$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!E1").Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M4").Value = "'.$data_input['generator_power_factor'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M5").Value = "'.($data_input['desired_coal_hhv']*4.1868).'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M6").Value = "'.$data_input['total_moisture'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M10").Value = "'.$data_input['excess_air'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M11").Value = "'.$data_input['ambient_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M12").Value = "'.$data_input['relative_humidity'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M13").Value = "'.$data_input['ambient_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M24").Value = "'.$data_input['water_cooling_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!M25").Value = "'.$data_input['cooling_temp_rise'].'"'."\r\n"; 

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
exec( $file_addr, $output, $retval);
$end_time = microtime(true);

$time_usage = ($end_time - $start_time);

//4. get output value
/** Load $inputFileName to a Spreadsheet object **/
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('processed_files\\'.$inputFileName);

$elink_status = $spreadsheet->getActiveSheet()->getCell('E6')->getValue();
// if($elink_status = 'Failed'){
// 	header("HTTP/1.1 500 Internal Server Error");
// 	echo 'Thermodynamic calculation failed.';
// 	die();
// }
//define output start
$generator_power_factor_value = round($spreadsheet->getSheetByName('output')->getCell('K5')->getCalculatedValue(),2);
$generator_power_factor_unit = $spreadsheet->getSheetByName('output')->getCell('H5')->getValue();
$gross_power_output_value = round($spreadsheet->getSheetByName('output')->getCell('K6')->getCalculatedValue(),2);
$gross_power_output_unit = $spreadsheet->getSheetByName('output')->getCell('H6')->getValue();
$auxiliary_power_consumption_value = round($spreadsheet->getSheetByName('output')->getCell('K7')->getCalculatedValue(),2);
$auxiliary_power_consumption_unit = $spreadsheet->getSheetByName('output')->getCell('H7')->getValue();
$desired_coal_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('K8')->getCalculatedValue(),2);
$desired_coal_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('H8')->getValue();
$gross_output_value = round($spreadsheet->getSheetByName('output')->getCell('K9')->getCalculatedValue(),2);
$gross_output_unit = $spreadsheet->getSheetByName('output')->getCell('H9')->getValue();
$boiler_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('K10')->getCalculatedValue(),2);
$boiler_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('H10')->getValue();
$net_output_value = round($spreadsheet->getSheetByName('output')->getCell('K11')->getCalculatedValue(),2);
$net_output_unit = $spreadsheet->getSheetByName('output')->getCell('H11')->getValue();
$net_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('K12')->getCalculatedValue(),2);
$net_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('H12')->getValue();
$gross_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('K13')->getCalculatedValue(),2);
$gross_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('H13')->getValue();
$plant_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('K14')->getCalculatedValue(),2);
$plant_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('H14')->getValue();
$ssh_steam_outlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K15')->getCalculatedValue(),2);
$ssh_steam_outlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H15')->getValue();
$ssh_steam_outlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K16')->getCalculatedValue(),2);
$ssh_steam_outlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H16')->getValue();
$ssh_steam_outlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K17')->getCalculatedValue(),2);
$ssh_steam_outlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H17')->getValue();
$main_steam_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K18')->getCalculatedValue(),2);
$main_steam_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H18')->getValue();
$main_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K19')->getCalculatedValue(),2);
$main_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H19')->getValue();
$main_steam_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K20')->getCalculatedValue(),2);
$main_steam_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H20')->getValue();
$lp_st_steam_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K21')->getCalculatedValue(),2);
$lp_st_steam_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H21')->getValue();
$lp_st_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K22')->getCalculatedValue(),2);
$lp_st_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H22')->getValue();
$lp_st_steam_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K23')->getCalculatedValue(),2);
$lp_st_steam_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H23')->getValue();
$cond_cw_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K24')->getCalculatedValue(),2);
$cond_cw_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H24')->getValue();
$hp_dry_step_eff1_value = round($spreadsheet->getSheetByName('output')->getCell('K25')->getCalculatedValue(),2);
$hp_dry_step_eff1_unit = $spreadsheet->getSheetByName('output')->getCell('H25')->getValue();
$hp_dry_step_eff2_value = round($spreadsheet->getSheetByName('output')->getCell('K26')->getCalculatedValue(),2);
$hp_dry_step_eff2_unit = $spreadsheet->getSheetByName('output')->getCell('H26')->getValue();
$hp_dry_step_eff3_value = round($spreadsheet->getSheetByName('output')->getCell('K27')->getCalculatedValue(),2);
$hp_dry_step_eff3_unit = $spreadsheet->getSheetByName('output')->getCell('H27')->getValue();
$hp_dry_step_eff4_value = round($spreadsheet->getSheetByName('output')->getCell('K28')->getCalculatedValue(),2);
$hp_dry_step_eff4_unit = $spreadsheet->getSheetByName('output')->getCell('H28')->getValue();
$hp_dry_step_eff5_value = round($spreadsheet->getSheetByName('output')->getCell('K29')->getCalculatedValue(),2);
$hp_dry_step_eff5_unit = $spreadsheet->getSheetByName('output')->getCell('H29')->getValue();
$lp_dry_step_eff1_value = round($spreadsheet->getSheetByName('output')->getCell('K30')->getCalculatedValue(),2);
$lp_dry_step_eff1_unit = $spreadsheet->getSheetByName('output')->getCell('H30')->getValue();
$lp_dry_step_eff2_value = round($spreadsheet->getSheetByName('output')->getCell('K31')->getCalculatedValue(),2);
$lp_dry_step_eff2_unit = $spreadsheet->getSheetByName('output')->getCell('H31')->getValue();
$lp_dry_step_eff3_value = round($spreadsheet->getSheetByName('output')->getCell('K32')->getCalculatedValue(),2);
$lp_dry_step_eff3_unit = $spreadsheet->getSheetByName('output')->getCell('H32')->getValue();
$fwh1_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K33')->getCalculatedValue(),2);
$fwh1_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H33')->getValue();
$fwh2_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K34')->getCalculatedValue(),2);
$fwh2_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H34')->getValue();
$fwh4_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K35')->getCalculatedValue(),2);
$fwh4_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H35')->getValue();
$fwh5_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K36')->getCalculatedValue(),2);
$fwh5_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H36')->getValue();
$fwh6_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K37')->getCalculatedValue(),2);
$fwh6_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H37')->getValue();
$fwh7_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('K38')->getCalculatedValue(),2);
$fwh7_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('H38')->getValue();
$fwh1_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K39')->getCalculatedValue(),2);
$fwh1_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H39')->getValue();
$fwh2_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K40')->getCalculatedValue(),2);
$fwh2_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H40')->getValue();
$fwh4_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K41')->getCalculatedValue(),2);
$fwh4_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H41')->getValue();
$fwh5_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K42')->getCalculatedValue(),2);
$fwh5_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H42')->getValue();
$fwh6_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K43')->getCalculatedValue(),2);
$fwh6_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H43')->getValue();
$fwh7_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('K44')->getCalculatedValue(),2);
$fwh7_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('H44')->getValue();
$cond_main_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K45')->getCalculatedValue(),2);
$cond_main_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H45')->getValue();
$cond_cooling_water_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K46')->getCalculatedValue(),2);
$cond_cooling_water_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H46')->getValue();
$cond_cooling_water_exit_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K47')->getCalculatedValue(),2);
$cond_cooling_water_exit_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H47')->getValue();
$cond_cooling_water_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K48')->getCalculatedValue(),2);
$cond_cooling_water_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H48')->getValue();
$coal_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K49')->getCalculatedValue(),2);
$coal_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H49')->getValue();
$stack_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K50')->getCalculatedValue(),2);
$stack_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H50')->getValue();
$aph_flue_gas_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K51')->getCalculatedValue(),2);
$aph_flue_gas_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H51')->getValue();
$sa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('K52')->getCalculatedValue(),2);
$sa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('H52')->getValue();
$pa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('K53')->getCalculatedValue(),2);
$pa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('H53')->getValue();
$flue_gas_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('K54')->getCalculatedValue(),2);
$flue_gas_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('H54')->getValue();
$sa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('K55')->getCalculatedValue(),2);
$sa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('H55')->getValue();
$pa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('K56')->getCalculatedValue(),2);
$pa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('H56')->getValue();
$lp_steam_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('K57')->getCalculatedValue(),2);
$lp_steam_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('H57')->getValue();
$hp_steam_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('K58')->getCalculatedValue(),2);
$hp_steam_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('H58')->getValue();
$ambient_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K59')->getCalculatedValue(),2);
$ambient_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H59')->getValue();
$relative_humidity_value = round($spreadsheet->getSheetByName('output')->getCell('K60')->getCalculatedValue(),2);
$relative_humidity_unit = $spreadsheet->getSheetByName('output')->getCell('H60')->getValue();
$ambient_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K61')->getCalculatedValue(),4);
$ambient_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H61')->getValue();
$feedwater_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('K62')->getCalculatedValue(),2);
$feedwater_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('H62')->getValue();
$feedwater_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('K63')->getCalculatedValue(),2);
$feedwater_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('H63')->getValue();
$feedwater_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('K64')->getCalculatedValue(),2);
$feedwater_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('H64')->getValue();



$json_subdata_int['gross_power_output']['value'] = round($gross_power_output_value/1000,2);
$json_subdata_int['gross_power_output']['unit'] = 'MW';
$json_subdata_int['net_output']['value'] = $net_output_value/1000;
$json_subdata_int['net_output']['unit'] = 'MW';
$json_subdata_int['net_heatrate']['value'] = round($net_heatrate_value/4.184,2);
$json_subdata_int['net_heatrate']['unit'] = 'kcal/kWh';//$net_heatrate_unit;
$json_subdata_int['gross_heatrate']['value'] = round($gross_heatrate_value/4.184,2);
$json_subdata_int['gross_heatrate']['unit'] = 'kcal/kWh';//$gross_heatrate_unit;
$json_subdata_int['time_usage']['value'] = round($time_usage/1000,3);
$json_subdata_int['time_usage']['unit'] = 'seconds';

$json_data['interface'] = $json_subdata_int;

$json_subdata_1['name']  = 'Generator Power Factor';
$json_subdata_1['unit'] = $generator_power_factor_unit;
$json_subdata_1['value'] = $generator_power_factor_value;
$json_subdata_2['name']  = 'Gross Power Output';
$json_subdata_2['unit'] = $gross_power_output_unit;
$json_subdata_2['value'] = $gross_power_output_value;
$json_subdata_3['name']  = 'Auxiliary Power consumption';
$json_subdata_3['unit'] = $auxiliary_power_consumption_unit;
$json_subdata_3['value'] = $auxiliary_power_consumption_value;
$json_subdata_4['name']  = 'Coal HHV';
$json_subdata_4['unit'] = $desired_coal_hhv_unit;
$json_subdata_4['value'] = $desired_coal_hhv_value;
$json_subdata_5['name']  = 'Gross Output';
$json_subdata_5['unit'] = $gross_output_unit;
$json_subdata_5['value'] = $gross_output_value;
$json_subdata_6['name']  = 'Boiler Efficiency';
$json_subdata_6['unit'] = $boiler_efficiency_unit;
$json_subdata_6['value'] = $boiler_efficiency_value;
$json_subdata_7['name']  = 'Net Output';
$json_subdata_7['unit'] = $net_output_unit;
$json_subdata_7['value'] = $net_output_value;
$json_subdata_8['name']  = 'Net Heatrate (HHV)';
$json_subdata_8['unit'] = $net_heatrate_unit;
$json_subdata_8['value'] = $net_heatrate_value;
$json_subdata_9['name']  = 'Gross Heatrate (LHV)';
$json_subdata_9['unit'] = $gross_heatrate_unit;
$json_subdata_9['value'] = $gross_heatrate_value;
$json_subdata_10['name']  = 'Plant Efficiency (HHV)';
$json_subdata_10['unit'] = $plant_efficiency_unit;
$json_subdata_10['value'] = $plant_efficiency_value;
$json_subdata_11['name']  = 'SSH Steam Outlet Temperature';
$json_subdata_11['unit'] = $ssh_steam_outlet_temperature_unit;
$json_subdata_11['value'] = $ssh_steam_outlet_temperature_value;
$json_subdata_12['name']  = 'SSH Steam Outlet Pressure';
$json_subdata_12['unit'] = $ssh_steam_outlet_pressure_unit;
$json_subdata_12['value'] = $ssh_steam_outlet_pressure_value;
$json_subdata_13['name']  = 'SSH Steam Outlet Flow';
$json_subdata_13['unit'] = $ssh_steam_outlet_flow_unit;
$json_subdata_13['value'] = $ssh_steam_outlet_flow_value;
$json_subdata_14['name']  = 'Main Steam Inlet Temperature';
$json_subdata_14['unit'] = $main_steam_inlet_temperature_unit;
$json_subdata_14['value'] = $main_steam_inlet_temperature_value;
$json_subdata_15['name']  = 'Main Steam Inlet Pressure';
$json_subdata_15['unit'] = $main_steam_inlet_pressure_unit;
$json_subdata_15['value'] = $main_steam_inlet_pressure_value;
$json_subdata_16['name']  = 'Main Steam Inlet Flow';
$json_subdata_16['unit'] = $main_steam_inlet_flow_unit;
$json_subdata_16['value'] = $main_steam_inlet_flow_value;
$json_subdata_17['name']  = 'LP ST Steam Inlet Temperature';
$json_subdata_17['unit'] = $lp_st_steam_inlet_temperature_unit;
$json_subdata_17['value'] = $lp_st_steam_inlet_temperature_value;
$json_subdata_18['name']  = 'LP ST Steam Inlet Pressure';
$json_subdata_18['unit'] = $lp_st_steam_inlet_pressure_unit;
$json_subdata_18['value'] = $lp_st_steam_inlet_pressure_value;
$json_subdata_19['name']  = 'LP ST Steam Inlet Flow';
$json_subdata_19['unit'] = $lp_st_steam_inlet_flow_unit;
$json_subdata_19['value'] = $lp_st_steam_inlet_flow_value;
$json_subdata_20['name']  = 'Cond. CW Flow';
$json_subdata_20['unit'] = $cond_cw_flow_unit;
$json_subdata_20['value'] = $cond_cw_flow_value;
$json_subdata_21['name']  = 'HP Dry Step Eff.';
$json_subdata_21['unit'] = $hp_dry_step_eff1_unit;
$json_subdata_21['value'] = $hp_dry_step_eff1_value;
$json_subdata_22['name']  = 'HP Dry Step Eff.';
$json_subdata_22['unit'] = $hp_dry_step_eff2_unit;
$json_subdata_22['value'] = $hp_dry_step_eff2_value;
$json_subdata_23['name']  = 'HP Dry Step Eff.';
$json_subdata_23['unit'] = $hp_dry_step_eff3_unit;
$json_subdata_23['value'] = $hp_dry_step_eff3_value;
$json_subdata_24['name']  = 'HP Dry Step Eff.';
$json_subdata_24['unit'] = $hp_dry_step_eff4_unit;
$json_subdata_24['value'] = $hp_dry_step_eff4_value;
$json_subdata_25['name']  = 'HP Dry Step Eff.';
$json_subdata_25['unit'] = $hp_dry_step_eff5_unit;
$json_subdata_25['value'] = $hp_dry_step_eff5_value;
$json_subdata_26['name']  = 'LP Dry Step Eff.';
$json_subdata_26['unit'] = $lp_dry_step_eff1_unit;
$json_subdata_26['value'] = $lp_dry_step_eff1_value;
$json_subdata_27['name']  = 'LP Dry Step Eff.';
$json_subdata_27['unit'] = $lp_dry_step_eff2_unit;
$json_subdata_27['value'] = $lp_dry_step_eff2_value;
$json_subdata_28['name']  = 'LP Dry Step Eff.';
$json_subdata_28['unit'] = $lp_dry_step_eff3_unit;
$json_subdata_28['value'] = $lp_dry_step_eff3_value;
$json_subdata_29['name']  = 'FWH1 Terminal Temperature Difference';
$json_subdata_29['unit'] = $fwh1_terminal_temperature_difference_unit;
$json_subdata_29['value'] = $fwh1_terminal_temperature_difference_value;
$json_subdata_30['name']  = 'FWH2 Terminal Temperature Difference';
$json_subdata_30['unit'] = $fwh2_terminal_temperature_difference_unit;
$json_subdata_30['value'] = $fwh2_terminal_temperature_difference_value;
$json_subdata_31['name']  = 'FWH4 Terminal Temperature Difference';
$json_subdata_31['unit'] = $fwh4_terminal_temperature_difference_unit;
$json_subdata_31['value'] = $fwh4_terminal_temperature_difference_value;
$json_subdata_32['name']  = 'FWH5 Terminal Temperature Difference';
$json_subdata_32['unit'] = $fwh5_terminal_temperature_difference_unit;
$json_subdata_32['value'] = $fwh5_terminal_temperature_difference_value;
$json_subdata_33['name']  = 'FWH6 Terminal Temperature Difference';
$json_subdata_33['unit'] = $fwh6_terminal_temperature_difference_unit;
$json_subdata_33['value'] = $fwh6_terminal_temperature_difference_value;
$json_subdata_34['name']  = 'FWH7 Terminal Temperature Difference';
$json_subdata_34['unit'] = $fwh7_terminal_temperature_difference_unit;
$json_subdata_34['value'] = $fwh7_terminal_temperature_difference_value;
$json_subdata_35['name']  = 'FWH1 Drain Cooler Approach';
$json_subdata_35['unit'] = $fwh1_drain_cooler_approach_unit;
$json_subdata_35['value'] = $fwh1_drain_cooler_approach_value;
$json_subdata_36['name']  = 'FWH2 Drain Cooler Approach';
$json_subdata_36['unit'] = $fwh2_drain_cooler_approach_unit;
$json_subdata_36['value'] = $fwh2_drain_cooler_approach_value;
$json_subdata_37['name']  = 'FWH4 Drain Cooler Approach';
$json_subdata_37['unit'] = $fwh4_drain_cooler_approach_unit;
$json_subdata_37['value'] = $fwh4_drain_cooler_approach_value;
$json_subdata_38['name']  = 'FWH5 Drain Cooler Approach';
$json_subdata_38['unit'] = $fwh5_drain_cooler_approach_unit;
$json_subdata_38['value'] = $fwh5_drain_cooler_approach_value;
$json_subdata_39['name']  = 'FWH6 Drain Cooler Approach';
$json_subdata_39['unit'] = $fwh6_drain_cooler_approach_unit;
$json_subdata_39['value'] = $fwh6_drain_cooler_approach_value;
$json_subdata_40['name']  = 'FWH7 Drain Cooler Approach';
$json_subdata_40['unit'] = $fwh7_drain_cooler_approach_unit;
$json_subdata_40['value'] = $fwh7_drain_cooler_approach_value;
$json_subdata_41['name']  = 'Cond. Main Steam Inlet Pressure';
$json_subdata_41['unit'] = $cond_main_steam_inlet_pressure_unit;
$json_subdata_41['value'] = $cond_main_steam_inlet_pressure_value;
$json_subdata_42['name']  = 'Cond. Cooling Water Inlet Temperature';
$json_subdata_42['unit'] = $cond_cooling_water_inlet_temperature_unit;
$json_subdata_42['value'] = $cond_cooling_water_inlet_temperature_value;
$json_subdata_43['name']  = 'Cond. Cooling Water Exit Temperature';
$json_subdata_43['unit'] = $cond_cooling_water_exit_temperature_unit;
$json_subdata_43['value'] = $cond_cooling_water_exit_temperature_value;
$json_subdata_44['name']  = 'Cond. Cooling Water Inlet Flow';
$json_subdata_44['unit'] = $cond_cooling_water_inlet_flow_unit;
$json_subdata_44['value'] = $cond_cooling_water_inlet_flow_value;
$json_subdata_45['name']  = 'Coal Flow';
$json_subdata_45['unit'] = $coal_flow_unit;
$json_subdata_45['value'] = $coal_flow_value;
$json_subdata_46['name']  = 'Stack Temperature';
$json_subdata_46['unit'] = $stack_temperature_unit;
$json_subdata_46['value'] = $stack_temperature_value;
$json_subdata_50['name']  = 'APH Flue gas inlet temperature';
$json_subdata_50['unit'] = $aph_flue_gas_inlet_temperature_unit;
$json_subdata_50['value'] = $aph_flue_gas_inlet_temperature_value;
$json_subdata_51['name']  = 'SA Outlet Temp.';
$json_subdata_51['unit'] = $sa_outlet_temp_unit;
$json_subdata_51['value'] = $sa_outlet_temp_value;
$json_subdata_52['name']  = 'PA Outlet Temp.';
$json_subdata_52['unit'] = $pa_outlet_temp_unit;
$json_subdata_52['value'] = $pa_outlet_temp_value;
$json_subdata_53['name']  = 'Flue Gas Outlet Temp';
$json_subdata_53['unit'] = $flue_gas_outlet_temp_unit;
$json_subdata_53['value'] = $flue_gas_outlet_temp_value;
$json_subdata_54['name']  = 'SA Inlet Temp';
$json_subdata_54['unit'] = $sa_inlet_temp_unit;
$json_subdata_54['value'] = $sa_inlet_temp_value;
$json_subdata_55['name']  = 'PA Inlet Temp';
$json_subdata_55['unit'] = $pa_inlet_temp_unit;
$json_subdata_55['value'] = $pa_inlet_temp_value;


$json_subdata_58['name']  = 'Ambient Temperature';
$json_subdata_58['unit'] = $ambient_temperature_unit;
$json_subdata_58['value'] = $ambient_temperature_value;
$json_subdata_59['name']  = 'Relative Humidity';
$json_subdata_59['unit'] = $relative_humidity_unit;
$json_subdata_59['value'] = $relative_humidity_value;
$json_subdata_60['name']  = 'Ambient Pressure';
$json_subdata_60['unit'] = $ambient_pressure_unit;
$json_subdata_60['value'] = $ambient_pressure_value;
$json_subdata_61['name']  = 'Feedwater mass flow';
$json_subdata_61['unit'] = $feedwater_mass_flow_unit;
$json_subdata_61['value'] = $feedwater_mass_flow_value;
$json_subdata_62['name']  = 'Feedwater Temperature';
$json_subdata_62['unit'] = $feedwater_temperature_unit;
$json_subdata_62['value'] = $feedwater_temperature_value;
$json_subdata_63['name']  = 'Feedwater Pressure';
$json_subdata_63['unit'] = $feedwater_pressure_unit;
$json_subdata_63['value'] = $feedwater_pressure_value;


$json_subdata_56['name']  = 'LP Steam Efficiency';
$json_subdata_56['unit'] = $lp_steam_efficiency_unit;
$json_subdata_56['value'] = $lp_steam_efficiency_value;
$json_subdata_57['name']  = 'HP Steam Efficiency';
$json_subdata_57['unit'] = $hp_steam_efficiency_unit;
$json_subdata_57['value'] = $hp_steam_efficiency_value;



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
$json_subdata_37,
$json_subdata_38,
$json_subdata_39,
$json_subdata_40,
$json_subdata_41,
$json_subdata_42,
$json_subdata_43,
$json_subdata_44,
$json_subdata_45,
$json_subdata_46,
$json_subdata_50,
$json_subdata_51,
$json_subdata_52,
$json_subdata_53,
$json_subdata_54,
$json_subdata_55,
$json_subdata_56,
$json_subdata_57,

$json_subdata_58,
$json_subdata_59,
$json_subdata_60,
$json_subdata_61,
$json_subdata_62,
$json_subdata_63
];


$json_draw_4['id'] = 'D31';
$json_draw_4['value'] = $desired_coal_hhv_value. ' '. $desired_coal_hhv_unit;
$json_draw_5['id'] = 'D53';
$json_draw_5['value'] = $gross_output_value. ' '. $gross_output_unit;
$json_draw_14['id'] = 'D41';
$json_draw_14['value'] = $main_steam_inlet_temperature_value. ' '. $main_steam_inlet_temperature_unit;
$json_draw_15['id'] = 'D42';
$json_draw_15['value'] = $main_steam_inlet_pressure_value. ' '. $main_steam_inlet_pressure_unit;
$json_draw_16['id'] = 'D40';
$json_draw_16['value'] = $main_steam_inlet_flow_value. ' '. $main_steam_inlet_flow_unit;
$json_draw_17['id'] = 'D51';
$json_draw_17['value'] = $lp_st_steam_inlet_temperature_value. ' '. $lp_st_steam_inlet_temperature_unit;
$json_draw_18['id'] = 'D52';
$json_draw_18['value'] = $lp_st_steam_inlet_pressure_value. ' '. $lp_st_steam_inlet_pressure_unit;
$json_draw_19['id'] = 'D50';
$json_draw_19['value'] = $lp_st_steam_inlet_flow_value. ' '. $lp_st_steam_inlet_flow_unit;
$json_draw_20['id'] = 'D20';
$json_draw_20['value'] = $cond_cw_flow_value. ' '. $cond_cw_flow_unit;
$json_draw_29['id'] = 'D11';
$json_draw_29['value'] = $fwh1_terminal_temperature_difference_value. ' '. $fwh1_terminal_temperature_difference_unit;
$json_draw_30['id'] = 'D09';
$json_draw_30['value'] = $fwh2_terminal_temperature_difference_value. ' '. $fwh2_terminal_temperature_difference_unit;
$json_draw_31['id'] = 'D07';
$json_draw_31['value'] = $fwh4_terminal_temperature_difference_value. ' '. $fwh4_terminal_temperature_difference_unit;
$json_draw_32['id'] = 'D05';
$json_draw_32['value'] = $fwh5_terminal_temperature_difference_value. ' '. $fwh5_terminal_temperature_difference_unit;
$json_draw_33['id'] = 'D03';
$json_draw_33['value'] = $fwh6_terminal_temperature_difference_value. ' '. $fwh6_terminal_temperature_difference_unit;
$json_draw_34['id'] = 'D01';
$json_draw_34['value'] = $fwh7_terminal_temperature_difference_value. ' '. $fwh7_terminal_temperature_difference_unit;
$json_draw_35['id'] = 'D12';
$json_draw_35['value'] = $fwh1_drain_cooler_approach_value. ' '. $fwh1_drain_cooler_approach_unit;
$json_draw_36['id'] = 'D10';
$json_draw_36['value'] = $fwh2_drain_cooler_approach_value. ' '. $fwh2_drain_cooler_approach_unit;
$json_draw_37['id'] = 'D08';
$json_draw_37['value'] = $fwh4_drain_cooler_approach_value. ' '. $fwh4_drain_cooler_approach_unit;
$json_draw_38['id'] = 'D06';
$json_draw_38['value'] = $fwh5_drain_cooler_approach_value. ' '. $fwh5_drain_cooler_approach_unit;
$json_draw_39['id'] = 'D04';
$json_draw_39['value'] = $fwh6_drain_cooler_approach_value. ' '. $fwh6_drain_cooler_approach_unit;
$json_draw_40['id'] = 'D02';
$json_draw_40['value'] = $fwh7_drain_cooler_approach_value. ' '. $fwh7_drain_cooler_approach_unit;
$json_draw_41['id'] = 'D18';
$json_draw_41['value'] = $cond_main_steam_inlet_pressure_value. ' '. $cond_main_steam_inlet_pressure_unit;
$json_draw_42['id'] = 'D19';
$json_draw_42['value'] = $cond_cooling_water_inlet_temperature_value. ' '. $cond_cooling_water_inlet_temperature_unit;
$json_draw_43['id'] = 'D21';
$json_draw_43['value'] = $cond_cooling_water_exit_temperature_value. ' '. $cond_cooling_water_exit_temperature_unit;
$json_draw_44['id'] = 'D20';
$json_draw_44['value'] = $cond_cooling_water_inlet_flow_value. ' '. $cond_cooling_water_inlet_flow_unit;
$json_draw_45['id'] = 'D30';
$json_draw_45['value'] = $coal_flow_value. ' '. $coal_flow_unit;
$json_draw_46['id'] = 'D39';
$json_draw_46['value'] = $stack_temperature_value. ' '. $stack_temperature_unit;
$json_draw_50['id'] = 'D32';
$json_draw_50['value'] = $aph_flue_gas_inlet_temperature_value. ' '. $aph_flue_gas_inlet_temperature_unit;
$json_draw_51['id'] = 'D36';
$json_draw_51['value'] = $sa_outlet_temp_value. ' '. $sa_outlet_temp_unit;
$json_draw_52['id'] = 'D34';
$json_draw_52['value'] = $pa_outlet_temp_value. ' '. $pa_outlet_temp_unit;
$json_draw_53['id'] = 'D33';
$json_draw_53['value'] = $flue_gas_outlet_temp_value. ' '. $flue_gas_outlet_temp_unit;
$json_draw_54['id'] = 'D37';
$json_draw_54['value'] = $sa_inlet_temp_value. ' '. $sa_inlet_temp_unit;
$json_draw_55['id'] = 'D35';
$json_draw_55['value'] = $pa_inlet_temp_value. ' '. $pa_inlet_temp_unit;
$json_draw_56['id'] = 'D56';
$json_draw_56['value'] = $lp_steam_efficiency_value. ' '. $lp_steam_efficiency_unit;
$json_draw_57['id'] = 'D57';
$json_draw_57['value'] = $hp_steam_efficiency_value. ' '. $hp_steam_efficiency_unit;
$json_draw_58['id'] = 'D61';
$json_draw_58['value'] = $ambient_temperature_value. ' '. $ambient_temperature_unit;
$json_draw_59['id'] = 'D62';
$json_draw_59['value'] = $relative_humidity_value. ' '. $relative_humidity_unit;
$json_draw_60['id'] = 'D63';
$json_draw_60['value'] = $ambient_pressure_value. ' '. $ambient_pressure_unit;
$json_draw_61['id'] = 'D15';
$json_draw_61['value'] = $feedwater_mass_flow_value. ' '. $feedwater_mass_flow_unit;
$json_draw_62['id'] = 'D16';
$json_draw_62['value'] = $feedwater_temperature_value. ' '. $feedwater_temperature_unit;
$json_draw_63['id'] = 'D17';
$json_draw_63['value'] = $feedwater_pressure_value. ' '. $feedwater_pressure_unit;


$json_data['drawing'] = [
	$json_draw_4,
$json_draw_5,
$json_draw_14,
$json_draw_15,
$json_draw_16,
$json_draw_17,
$json_draw_18,
$json_draw_19,
$json_draw_20,
$json_draw_29,
$json_draw_30,
$json_draw_31,
$json_draw_32,
$json_draw_33,
$json_draw_34,
$json_draw_35,
$json_draw_36,
$json_draw_37,
$json_draw_38,
$json_draw_39,
$json_draw_40,
$json_draw_41,
$json_draw_42,
$json_draw_43,
$json_draw_44,
$json_draw_45,
$json_draw_46,
$json_draw_50,
$json_draw_51,
$json_draw_52,
$json_draw_53,
$json_draw_54,
$json_draw_55,
$json_draw_56,
$json_draw_57,

$json_draw_58,
$json_draw_59,
$json_draw_60,
$json_draw_61,
$json_draw_62,
$json_draw_63
];

//5. serve output value
echo json_encode($json_data);
} catch (Exception $e){
	header("HTTP/1.1 500 Internal Server Error");
	echo $e->getMessage();
	die();
}
?>