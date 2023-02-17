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

$fileRep = 'tja_master'.((float)$usec + (float)$sec);
$fileRep = str_replace(".","",$fileRep);
$inputFileName = $fileRep.'.xlsm';
copy("tja_master.xlsm",'processed_files\\'.$inputFileName);

//2. write value to file
//generate file
$vbs_content = 'Set objExcel_'.$fileRep.'    = CreateObject("Excel.Application")'."\r\n";
$vbs_content = $vbs_content. 'Set objWorkbook_'.$fileRep.' = objExcel_'.$fileRep.'.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\\'.$inputFileName.'")'."\r\n";

// set value input start
//Worksheets("Input").Activate
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!E1").Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N3").Value = "'.$data_input['generator_power_factor'].'"'."\r\n"; 
// $vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(238,5).Value = "'.$data_input['expected_power_output'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N4").Value = "'.($data_input['desired_coal_hhv']*4.1868).'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N5").Value = "'.$data_input['total_moisture'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N12").Value = "'.$data_input['ambient_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N13").Value = "'.$data_input['relative_humidity'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N14").Value = "'.$data_input['ambient_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N11").Value = "'.$data_input['excess_air'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N26").Value = "'.$data_input['water_cooling_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Range("input!N27").Value = "'.$data_input['cooling_temp_rise'].'"'."\r\n"; 


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

// $elink_status = $spreadsheet->getSheetByName('output')->getCell('E6')->getCalculatedValue();
// if($elink_status = 'Failed'){
// 	header("HTTP/1.1 500 Internal Server Error");
// 	echo 'Thermodynamic calculation failed.';
// 	die();
// }

$spreadsheet->setActiveSheetIndex(5);

//defaine output start
$generator_power_factor_value = round($spreadsheet->getSheetByName('output')->getCell('L5')->getCalculatedValue(),2);
$generator_power_factor_unit = $spreadsheet->getSheetByName('output')->getCell('I5')->getCalculatedValue();
$expected_power_output_value = round($spreadsheet->getSheetByName('output')->getCell('L6')->getCalculatedValue(),2);
$expected_power_output_unit = $spreadsheet->getSheetByName('output')->getCell('I6')->getCalculatedValue();
$auxiliary_power_consumption_value = round($spreadsheet->getSheetByName('output')->getCell('L7')->getCalculatedValue(),2);
$auxiliary_power_consumption_unit = $spreadsheet->getSheetByName('output')->getCell('I7')->getCalculatedValue();
$desired_coal_hhv_value = round($spreadsheet->getSheetByName('output')->getCell('L8')->getCalculatedValue(),2);
$desired_coal_hhv_unit = $spreadsheet->getSheetByName('output')->getCell('I8')->getCalculatedValue();
$gross_output_value = round($spreadsheet->getSheetByName('output')->getCell('L9')->getCalculatedValue(),2);
$gross_output_unit = $spreadsheet->getSheetByName('output')->getCell('I9')->getCalculatedValue();
$boiler_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('L10')->getCalculatedValue(),2);
$boiler_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('I10')->getCalculatedValue();
$net_output_value = round($spreadsheet->getSheetByName('output')->getCell('L11')->getCalculatedValue(),2);
$net_output_unit = $spreadsheet->getSheetByName('output')->getCell('I11')->getCalculatedValue();
$net_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('L12')->getCalculatedValue(),2);
$net_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('I12')->getCalculatedValue();
$gross_heatrate_value = round($spreadsheet->getSheetByName('output')->getCell('L13')->getCalculatedValue(),2);
$gross_heatrate_unit = $spreadsheet->getSheetByName('output')->getCell('I13')->getCalculatedValue();
$plant_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('L14')->getCalculatedValue(),2);
$plant_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('I14')->getCalculatedValue();
$ssh_steam_outlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L15')->getCalculatedValue(),2);
$ssh_steam_outlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I15')->getCalculatedValue();
$ssh_steam_outlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L16')->getCalculatedValue(),2);
$ssh_steam_outlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I16')->getCalculatedValue();
$ssh_steam_outlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L17')->getCalculatedValue(),2);
$ssh_steam_outlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I17')->getCalculatedValue();
$frh_steam_outlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L18')->getCalculatedValue(),2);
$frh_steam_outlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I18')->getCalculatedValue();
$frh_steam_outlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L19')->getCalculatedValue(),2);
$frh_steam_outlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I19')->getCalculatedValue();
$frh_steam_outlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L20')->getCalculatedValue(),2);
$frh_steam_outlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I20')->getCalculatedValue();
$main_steam_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L21')->getCalculatedValue(),2);
$main_steam_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I21')->getCalculatedValue();
$main_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L22')->getCalculatedValue(),2);
$main_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I22')->getCalculatedValue();
$main_steam_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L23')->getCalculatedValue(),2);
$main_steam_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I23')->getCalculatedValue();
$hrh_steam_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L24')->getCalculatedValue(),2);
$hrh_steam_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I24')->getCalculatedValue();
$hrh_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L25')->getCalculatedValue(),2);
$hrh_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I25')->getCalculatedValue();
$hrh_steam_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L26')->getCalculatedValue(),2);
$hrh_steam_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I26')->getCalculatedValue();
$lp_st_steam_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L27')->getCalculatedValue(),2);
$lp_st_steam_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I27')->getCalculatedValue();
$lp_st_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L28')->getCalculatedValue(),2);
$lp_st_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I28')->getCalculatedValue();
$lp_st_steam_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L29')->getCalculatedValue(),2);
$lp_st_steam_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I29')->getCalculatedValue();
$cond_cw_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L30')->getCalculatedValue(),2);
$cond_cw_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I30')->getCalculatedValue();
$hp_dry_step_eff1_value = round($spreadsheet->getSheetByName('output')->getCell('L31')->getCalculatedValue(),2);
$hp_dry_step_eff1_unit = $spreadsheet->getSheetByName('output')->getCell('I31')->getCalculatedValue();
$hp_dry_step_eff2_value = round($spreadsheet->getSheetByName('output')->getCell('L32')->getCalculatedValue(),2);
$hp_dry_step_eff2_unit = $spreadsheet->getSheetByName('output')->getCell('I32')->getCalculatedValue();
$ip_dry_step_eff1_value = round($spreadsheet->getSheetByName('output')->getCell('L33')->getCalculatedValue(),2);
$ip_dry_step_eff1_unit = $spreadsheet->getSheetByName('output')->getCell('I33')->getCalculatedValue();
$ip_dry_step_eff2_value = round($spreadsheet->getSheetByName('output')->getCell('L34')->getCalculatedValue(),2);
$ip_dry_step_eff2_unit = $spreadsheet->getSheetByName('output')->getCell('I34')->getCalculatedValue();
$lp_dry_step_eff1_value = round($spreadsheet->getSheetByName('output')->getCell('L35')->getCalculatedValue(),2);
$lp_dry_step_eff1_unit = $spreadsheet->getSheetByName('output')->getCell('I35')->getCalculatedValue();
$lp_dry_step_eff2_value = round($spreadsheet->getSheetByName('output')->getCell('L36')->getCalculatedValue(),2);
$lp_dry_step_eff2_unit = $spreadsheet->getSheetByName('output')->getCell('I36')->getCalculatedValue();
$lp_dry_step_eff3_value = round($spreadsheet->getSheetByName('output')->getCell('L37')->getCalculatedValue(),2);
$lp_dry_step_eff3_unit = $spreadsheet->getSheetByName('output')->getCell('I37')->getCalculatedValue();
$lp_dry_step_eff4_value = round($spreadsheet->getSheetByName('output')->getCell('L38')->getCalculatedValue(),2);
$lp_dry_step_eff4_unit = $spreadsheet->getSheetByName('output')->getCell('I38')->getCalculatedValue();
$lp_dry_step_eff5_value = round($spreadsheet->getSheetByName('output')->getCell('L39')->getCalculatedValue(),2);
$lp_dry_step_eff5_unit = $spreadsheet->getSheetByName('output')->getCell('I39')->getCalculatedValue();
$fwh1_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L40')->getCalculatedValue(),2);
$fwh1_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I40')->getCalculatedValue();
$fwh2_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L41')->getCalculatedValue(),2);
$fwh2_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I41')->getCalculatedValue();
$fwh3_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L42')->getCalculatedValue(),2);
$fwh3_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I42')->getCalculatedValue();
$fwh5_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L43')->getCalculatedValue(),2);
$fwh5_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I43')->getCalculatedValue();
$fwh6_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L44')->getCalculatedValue(),2);
$fwh6_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I44')->getCalculatedValue();
$fwh7_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L45')->getCalculatedValue(),2);
$fwh7_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I45')->getCalculatedValue();
$fwh8_terminal_temperature_difference_value = round($spreadsheet->getSheetByName('output')->getCell('L46')->getCalculatedValue(),2);
$fwh8_terminal_temperature_difference_unit = $spreadsheet->getSheetByName('output')->getCell('I46')->getCalculatedValue();
$fwh1_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L47')->getCalculatedValue(),2);
$fwh1_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I47')->getCalculatedValue();
$fwh2_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L48')->getCalculatedValue(),2);
$fwh2_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I48')->getCalculatedValue();
$fwh3_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L49')->getCalculatedValue(),2);
$fwh3_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I49')->getCalculatedValue();
$fwh5_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L50')->getCalculatedValue(),2);
$fwh5_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I50')->getCalculatedValue();
$fwh6_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L51')->getCalculatedValue(),2);
$fwh6_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I51')->getCalculatedValue();
$fwh7_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L52')->getCalculatedValue(),2);
$fwh7_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I52')->getCalculatedValue();
$fwh8_drain_cooler_approach_value = round($spreadsheet->getSheetByName('output')->getCell('L53')->getCalculatedValue(),2);
$fwh8_drain_cooler_approach_unit = $spreadsheet->getSheetByName('output')->getCell('I53')->getCalculatedValue();
$cond_main_steam_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L54')->getCalculatedValue(),2);
$cond_main_steam_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I54')->getCalculatedValue();

$cond_cooling_water_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L56')->getCalculatedValue(),2);
$cond_cooling_water_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I56')->getCalculatedValue();
$cond_cooling_water_exit_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L57')->getCalculatedValue(),2);
$cond_cooling_water_exit_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I57')->getCalculatedValue();
$cond_cooling_water_inlet_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L58')->getCalculatedValue(),2);
$cond_cooling_water_inlet_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I58')->getCalculatedValue();
$coal_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L59')->getCalculatedValue(),2);
$coal_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I59')->getCalculatedValue();
$stack_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L60')->getCalculatedValue(),2);
$stack_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I60')->getCalculatedValue();
$ambient_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L61')->getCalculatedValue(),2);
$ambient_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I61')->getCalculatedValue();
$ambient_rh_value = round($spreadsheet->getSheetByName('output')->getCell('L62')->getCalculatedValue(),2);
$ambient_rh_unit = $spreadsheet->getSheetByName('output')->getCell('I62')->getCalculatedValue();
$cond_condensate_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L63')->getCalculatedValue(),2);
$cond_condensate_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I63')->getCalculatedValue();
$aph_flue_gas_inlet_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L64')->getCalculatedValue(),2);
$aph_flue_gas_inlet_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I64')->getCalculatedValue();
$sa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L65')->getCalculatedValue(),2);
$sa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I65')->getCalculatedValue();
$pa_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L66')->getCalculatedValue(),2);
$pa_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I66')->getCalculatedValue();
$flue_gas_outlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L67')->getCalculatedValue(),2);
$flue_gas_outlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I67')->getCalculatedValue();
$sa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L68')->getCalculatedValue(),2);
$sa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I68')->getCalculatedValue();
$pa_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L69')->getCalculatedValue(),2);
$pa_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I69')->getCalculatedValue();
$crh_inlet_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L70')->getCalculatedValue(),2);
$crh_inlet_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I70')->getCalculatedValue();
$crh_inlet_temp_value = round($spreadsheet->getSheetByName('output')->getCell('L71')->getCalculatedValue(),2);
$crh_inlet_temp_unit = $spreadsheet->getSheetByName('output')->getCell('I71')->getCalculatedValue();
$crh_inlet_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L72')->getCalculatedValue(),2);
$crh_inlet_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I72')->getCalculatedValue();
$lp_steam_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('L73')->getCalculatedValue(),2);
$lp_steam_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('I73')->getCalculatedValue();
$ip_steam_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('L74')->getCalculatedValue(),2);
$ip_steam_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('I74')->getCalculatedValue();
$hp_steam_efficiency_value = round($spreadsheet->getSheetByName('output')->getCell('L75')->getCalculatedValue(),2);
$hp_steam_efficiency_unit = $spreadsheet->getSheetByName('output')->getCell('I75')->getCalculatedValue();
$ambient_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L76')->getCalculatedValue(),2);
$ambient_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I76')->getCalculatedValue();
$relative_humidity_value = round($spreadsheet->getSheetByName('output')->getCell('L77')->getCalculatedValue(),2);
$relative_humidity_unit = $spreadsheet->getSheetByName('output')->getCell('I77')->getCalculatedValue();
$ambient_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L78')->getCalculatedValue(),2);
$ambient_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I78')->getCalculatedValue();
$feedwater_mass_flow_value = round($spreadsheet->getSheetByName('output')->getCell('L79')->getCalculatedValue(),2);
$feedwater_mass_flow_unit = $spreadsheet->getSheetByName('output')->getCell('I79')->getCalculatedValue();
$feedwater_temperature_value = round($spreadsheet->getSheetByName('output')->getCell('L80')->getCalculatedValue(),2);
$feedwater_temperature_unit = $spreadsheet->getSheetByName('output')->getCell('I80')->getCalculatedValue();
$feedwater_pressure_value = round($spreadsheet->getSheetByName('output')->getCell('L81')->getCalculatedValue(),2);
$feedwater_pressure_unit = $spreadsheet->getSheetByName('output')->getCell('I81')->getCalculatedValue();


$json_subdata_int['gross_power_output']['value'] = round($gross_output_value/1000,2);
$json_subdata_int['gross_power_output']['unit'] = 'MW';
$json_subdata_int['net_output']['value'] = round($net_output_value/1000,2);
$json_subdata_int['net_output']['unit'] = 'MW';
$json_subdata_int['net_heatrate']['value'] = round($net_heatrate_value,2);
$json_subdata_int['net_heatrate']['unit'] = $net_heatrate_unit;
$json_subdata_int['gross_heatrate']['value'] = round($gross_heatrate_value,2);
$json_subdata_int['gross_heatrate']['unit'] = $gross_heatrate_unit;
$json_subdata_int['time_usage']['value'] = round($time_usage,3);
$json_subdata_int['time_usage']['unit'] = 'seconds';


$json_data['interface'] = $json_subdata_int;

$json_subdata_1['name'] = 'Generator Power Factor';
$json_subdata_1['unit'] = $generator_power_factor_unit;
$json_subdata_1['value'] = $generator_power_factor_value;
$json_subdata_2['name'] = 'Gross Power Output';
$json_subdata_2['unit'] = $gross_output_unit;
$json_subdata_2['value'] = $gross_output_value;
$json_subdata_3['name'] = 'Auxiliary Power consumption';
$json_subdata_3['unit'] = $auxiliary_power_consumption_unit;
$json_subdata_3['value'] = $auxiliary_power_consumption_value;
$json_subdata_4['name'] = 'Coal HHV';
$json_subdata_4['unit'] = $desired_coal_hhv_unit;
$json_subdata_4['value'] = $desired_coal_hhv_value;
$json_subdata_5['name'] = 'Gross Output';
$json_subdata_5['unit'] = $gross_output_unit;
$json_subdata_5['value'] = $gross_output_value;
$json_subdata_6['name'] = 'Boiler Efficiency';
$json_subdata_6['unit'] = $boiler_efficiency_unit;
$json_subdata_6['value'] = $boiler_efficiency_value;
$json_subdata_7['name'] = 'Net Output';
$json_subdata_7['unit'] = $net_output_unit;
$json_subdata_7['value'] = $net_output_value;
$json_subdata_8['name'] = 'Net Heatrate (HHV)';
$json_subdata_8['unit'] = $net_heatrate_unit;
$json_subdata_8['value'] = $net_heatrate_value;
$json_subdata_9['name'] = 'Gross Heatrate (LHV)';
$json_subdata_9['unit'] = $gross_heatrate_unit;
$json_subdata_9['value'] = $gross_heatrate_value;
$json_subdata_10['name'] = 'Plant Efficiency (HHV)';
$json_subdata_10['unit'] = $plant_efficiency_unit;
$json_subdata_10['value'] = $plant_efficiency_value;
$json_subdata_11['name'] = 'SSH Steam Outlet Temperature';
$json_subdata_11['unit'] = $ssh_steam_outlet_temperature_unit;
$json_subdata_11['value'] = $ssh_steam_outlet_temperature_value;
$json_subdata_12['name'] = 'SSH Steam Outlet Pressure';
$json_subdata_12['unit'] = $ssh_steam_outlet_pressure_unit;
$json_subdata_12['value'] = $ssh_steam_outlet_pressure_value;
$json_subdata_13['name'] = 'SSH Steam Outlet Flow';
$json_subdata_13['unit'] = $ssh_steam_outlet_flow_unit;
$json_subdata_13['value'] = $ssh_steam_outlet_flow_value;
$json_subdata_14['name'] = 'FRH Steam Outlet Temperature';
$json_subdata_14['unit'] = $frh_steam_outlet_temperature_unit;
$json_subdata_14['value'] = $frh_steam_outlet_temperature_value;
$json_subdata_15['name'] = 'FRH Steam Outlet Pressure';
$json_subdata_15['unit'] = $frh_steam_outlet_pressure_unit;
$json_subdata_15['value'] = $frh_steam_outlet_pressure_value;
$json_subdata_16['name'] = 'FRH Steam Outlet Flow';
$json_subdata_16['unit'] = $frh_steam_outlet_flow_unit;
$json_subdata_16['value'] = $frh_steam_outlet_flow_value;
$json_subdata_17['name'] = 'Main Steam Inlet Temperature';
$json_subdata_17['unit'] = $main_steam_inlet_temperature_unit;
$json_subdata_17['value'] = $main_steam_inlet_temperature_value;
$json_subdata_18['name'] = 'Main Steam Inlet Pressure';
$json_subdata_18['unit'] = $main_steam_inlet_pressure_unit;
$json_subdata_18['value'] = $main_steam_inlet_pressure_value;
$json_subdata_19['name'] = 'Main Steam Inlet Flow';
$json_subdata_19['unit'] = $main_steam_inlet_flow_unit;
$json_subdata_19['value'] = $main_steam_inlet_flow_value;
$json_subdata_20['name'] = 'HRH Steam Inlet Temperature';
$json_subdata_20['unit'] = $hrh_steam_inlet_temperature_unit;
$json_subdata_20['value'] = $hrh_steam_inlet_temperature_value;
$json_subdata_21['name'] = 'HRH Steam Inlet Pressure';
$json_subdata_21['unit'] = $hrh_steam_inlet_pressure_unit;
$json_subdata_21['value'] = $hrh_steam_inlet_pressure_value;
$json_subdata_22['name'] = 'HRH Steam Inlet Flow';
$json_subdata_22['unit'] = $hrh_steam_inlet_flow_unit;
$json_subdata_22['value'] = $hrh_steam_inlet_flow_value;
$json_subdata_23['name'] = 'LP ST Steam Inlet Temperature';
$json_subdata_23['unit'] = $lp_st_steam_inlet_temperature_unit;
$json_subdata_23['value'] = $lp_st_steam_inlet_temperature_value;
$json_subdata_24['name'] = 'LP ST Steam Inlet Pressure';
$json_subdata_24['unit'] = $lp_st_steam_inlet_pressure_unit;
$json_subdata_24['value'] = $lp_st_steam_inlet_pressure_value;
$json_subdata_25['name'] = 'LP ST Steam Inlet Flow';
$json_subdata_25['unit'] = $lp_st_steam_inlet_flow_unit;
$json_subdata_25['value'] = $lp_st_steam_inlet_flow_value;
$json_subdata_26['name'] = 'Cond. CW Flow';
$json_subdata_26['unit'] = $cond_cw_flow_unit;
$json_subdata_26['value'] = $cond_cw_flow_value;
$json_subdata_27['name'] = 'HP Dry Step Eff.';
$json_subdata_27['unit'] = $hp_dry_step_eff1_unit;
$json_subdata_27['value'] = $hp_dry_step_eff1_value;
$json_subdata_28['name'] = 'HP Dry Step Eff.';
$json_subdata_28['unit'] = $hp_dry_step_eff2_unit;
$json_subdata_28['value'] = $hp_dry_step_eff2_value;
$json_subdata_29['name'] = 'IP Dry Step Eff.';
$json_subdata_29['unit'] = $ip_dry_step_eff1_unit;
$json_subdata_29['value'] = $ip_dry_step_eff1_value;
$json_subdata_30['name'] = 'IP Dry Step Eff.';
$json_subdata_30['unit'] = $ip_dry_step_eff2_unit;
$json_subdata_30['value'] = $ip_dry_step_eff2_value;
$json_subdata_31['name'] = 'LP Dry Step Eff.';
$json_subdata_31['unit'] = $lp_dry_step_eff1_unit;
$json_subdata_31['value'] = $lp_dry_step_eff1_value;
$json_subdata_32['name'] = 'LP Dry Step Eff.';
$json_subdata_32['unit'] = $lp_dry_step_eff2_unit;
$json_subdata_32['value'] = $lp_dry_step_eff2_value;
$json_subdata_33['name'] = 'LP Dry Step Eff.';
$json_subdata_33['unit'] = $lp_dry_step_eff3_unit;
$json_subdata_33['value'] = $lp_dry_step_eff3_value;
$json_subdata_34['name'] = 'LP Dry Step Eff.';
$json_subdata_34['unit'] = $lp_dry_step_eff4_unit;
$json_subdata_34['value'] = $lp_dry_step_eff4_value;
$json_subdata_35['name'] = 'LP Dry Step Eff.';
$json_subdata_35['unit'] = $lp_dry_step_eff5_unit;
$json_subdata_35['value'] = $lp_dry_step_eff5_value;
$json_subdata_36['name'] = 'FWH1 Terminal Temperature Difference';
$json_subdata_36['unit'] = $fwh1_terminal_temperature_difference_unit;
$json_subdata_36['value'] = $fwh1_terminal_temperature_difference_value;
$json_subdata_37['name'] = 'FWH2 Terminal Temperature Difference';
$json_subdata_37['unit'] = $fwh2_terminal_temperature_difference_unit;
$json_subdata_37['value'] = $fwh2_terminal_temperature_difference_value;
$json_subdata_38['name'] = 'FWH3 Terminal Temperature Difference';
$json_subdata_38['unit'] = $fwh3_terminal_temperature_difference_unit;
$json_subdata_38['value'] = $fwh3_terminal_temperature_difference_value;
$json_subdata_39['name'] = 'FWH5 Terminal Temperature Difference';
$json_subdata_39['unit'] = $fwh5_terminal_temperature_difference_unit;
$json_subdata_39['value'] = $fwh5_terminal_temperature_difference_value;
$json_subdata_40['name'] = 'FWH6 Terminal Temperature Difference';
$json_subdata_40['unit'] = $fwh6_terminal_temperature_difference_unit;
$json_subdata_40['value'] = $fwh6_terminal_temperature_difference_value;
$json_subdata_41['name'] = 'FWH7 Terminal Temperature Difference';
$json_subdata_41['unit'] = $fwh7_terminal_temperature_difference_unit;
$json_subdata_41['value'] = $fwh7_terminal_temperature_difference_value;
$json_subdata_42['name'] = 'FWH8 Terminal Temperature Difference';
$json_subdata_42['unit'] = $fwh8_terminal_temperature_difference_unit;
$json_subdata_42['value'] = $fwh8_terminal_temperature_difference_value;
$json_subdata_43['name'] = 'FWH1 Drain Cooler Approach';
$json_subdata_43['unit'] = $fwh1_drain_cooler_approach_unit;
$json_subdata_43['value'] = $fwh1_drain_cooler_approach_value;
$json_subdata_44['name'] = 'FWH2 Drain Cooler Approach';
$json_subdata_44['unit'] = $fwh2_drain_cooler_approach_unit;
$json_subdata_44['value'] = $fwh2_drain_cooler_approach_value;
$json_subdata_45['name'] = 'FWH3 Drain Cooler Approach';
$json_subdata_45['unit'] = $fwh3_drain_cooler_approach_unit;
$json_subdata_45['value'] = $fwh3_drain_cooler_approach_value;
$json_subdata_46['name'] = 'FWH5 Drain Cooler Approach';
$json_subdata_46['unit'] = $fwh5_drain_cooler_approach_unit;
$json_subdata_46['value'] = $fwh5_drain_cooler_approach_value;
$json_subdata_47['name'] = 'FWH6 Drain Cooler Approach';
$json_subdata_47['unit'] = $fwh6_drain_cooler_approach_unit;
$json_subdata_47['value'] = $fwh6_drain_cooler_approach_value;
$json_subdata_48['name'] = 'FWH7 Drain Cooler Approach';
$json_subdata_48['unit'] = $fwh7_drain_cooler_approach_unit;
$json_subdata_48['value'] = $fwh7_drain_cooler_approach_value;
$json_subdata_49['name'] = 'FWH8 Drain Cooler Approach';
$json_subdata_49['unit'] = $fwh8_drain_cooler_approach_unit;
$json_subdata_49['value'] = $fwh8_drain_cooler_approach_value;
$json_subdata_50['name'] = 'Cond. Main Steam Inlet Pressure';
$json_subdata_50['unit'] = $cond_main_steam_inlet_pressure_unit;
$json_subdata_50['value'] = $cond_main_steam_inlet_pressure_value;

$json_subdata_52['name'] = 'Cond. Cooling Water Inlet Temperature';
$json_subdata_52['unit'] = $cond_cooling_water_inlet_temperature_unit;
$json_subdata_52['value'] = $cond_cooling_water_inlet_temperature_value;
$json_subdata_53['name'] = 'Cond. Cooling Water Exit Temperature';
$json_subdata_53['unit'] = $cond_cooling_water_exit_temperature_unit;
$json_subdata_53['value'] = $cond_cooling_water_exit_temperature_value;
$json_subdata_54['name'] = 'Cond. Cooling Water Inlet Flow';
$json_subdata_54['unit'] = $cond_cooling_water_inlet_flow_unit;
$json_subdata_54['value'] = $cond_cooling_water_inlet_flow_value;



$json_subdata_58['name'] = 'Coal Flow';
$json_subdata_58['unit'] = $coal_flow_unit;
$json_subdata_58['value'] = $coal_flow_value;
$json_subdata_59['name'] = 'Stack Temperature';
$json_subdata_59['unit'] = $stack_temperature_unit;
$json_subdata_59['value'] = $stack_temperature_value;
$json_subdata_60['name'] = 'Ambient Pressure';
$json_subdata_60['unit'] = $ambient_pressure_unit;
$json_subdata_60['value'] = $ambient_pressure_value;
$json_subdata_61['name'] = 'Ambient RH';
$json_subdata_61['unit'] = $ambient_rh_unit;
$json_subdata_61['value'] = $ambient_rh_value;
$json_subdata_62['name'] = 'Cond. Condensate temperature';
$json_subdata_62['unit'] = $cond_condensate_temperature_unit;
$json_subdata_62['value'] = $cond_condensate_temperature_value;
$json_subdata_63['name'] = 'APH Flue gas inlet temperature';
$json_subdata_63['unit'] = $aph_flue_gas_inlet_temperature_unit;
$json_subdata_63['value'] = $aph_flue_gas_inlet_temperature_value;
$json_subdata_64['name'] = 'SA Outlet Temp.';
$json_subdata_64['unit'] = $sa_outlet_temp_unit;
$json_subdata_64['value'] = $sa_outlet_temp_value;
$json_subdata_65['name'] = 'PA Outlet Temp.';
$json_subdata_65['unit'] = $pa_outlet_temp_unit;
$json_subdata_65['value'] = $pa_outlet_temp_value;
$json_subdata_66['name'] = 'Flue Gas Outlet Temp';
$json_subdata_66['unit'] = $flue_gas_outlet_temp_unit;
$json_subdata_66['value'] = $flue_gas_outlet_temp_value;
$json_subdata_67['name'] = 'SA Inlet Temp';
$json_subdata_67['unit'] = $sa_inlet_temp_unit;
$json_subdata_67['value'] = $sa_inlet_temp_value;
$json_subdata_68['name'] = 'PA Inlet Temp';
$json_subdata_68['unit'] = $pa_inlet_temp_unit;
$json_subdata_68['value'] = $pa_inlet_temp_value;
$json_subdata_69['name'] = 'CRH Inlet Mass Flow';
$json_subdata_69['unit'] = $crh_inlet_mass_flow_unit;
$json_subdata_69['value'] = $crh_inlet_mass_flow_value;
$json_subdata_70['name'] = 'CRH Inlet Temp';
$json_subdata_70['unit'] = $crh_inlet_temp_unit;
$json_subdata_70['value'] = $crh_inlet_temp_value;
$json_subdata_71['name'] = 'CRH Inlet Pressure';
$json_subdata_71['unit'] = $crh_inlet_pressure_unit;
$json_subdata_71['value'] = $crh_inlet_pressure_value;
$json_subdata_72['name'] = 'LP Steam Efficiency';
$json_subdata_72['unit'] = $lp_steam_efficiency_unit;
$json_subdata_72['value'] = $lp_steam_efficiency_value;
$json_subdata_73['name'] = 'IP Steam Efficiency';
$json_subdata_73['unit'] = $ip_steam_efficiency_unit;
$json_subdata_73['value'] = $ip_steam_efficiency_value;
$json_subdata_74['name'] = 'HP Steam Efficiency';
$json_subdata_74['unit'] = $hp_steam_efficiency_unit;
$json_subdata_74['value'] = $hp_steam_efficiency_value;
$json_subdata_75['name'] = 'Ambient Temperature';
$json_subdata_75['unit'] = $ambient_temperature_unit;
$json_subdata_75['value'] = $ambient_temperature_value;
$json_subdata_76['name'] = 'Relative Humidity';
$json_subdata_76['unit'] = $relative_humidity_unit;
$json_subdata_76['value'] = $relative_humidity_value;
$json_subdata_77['name'] = 'Ambient Pressure';
$json_subdata_77['unit'] = $ambient_pressure_unit;
$json_subdata_77['value'] = $ambient_pressure_value;
$json_subdata_78['name'] = 'Feedwater mass flow';
$json_subdata_78['unit'] = $feedwater_mass_flow_unit;
$json_subdata_78['value'] = $feedwater_mass_flow_value;
$json_subdata_79['name'] = 'Feedwater Temperature';
$json_subdata_79['unit'] = $feedwater_temperature_unit;
$json_subdata_79['value'] = $feedwater_temperature_value;
$json_subdata_80['name'] = 'Feedwater Pressure';
$json_subdata_80['unit'] = $feedwater_pressure_unit;
$json_subdata_80['value'] = $feedwater_pressure_value;




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
$json_subdata_47,
$json_subdata_48,
$json_subdata_49,
$json_subdata_50,

$json_subdata_52,
$json_subdata_53,
$json_subdata_54,

$json_subdata_58,
$json_subdata_59,
$json_subdata_60,
$json_subdata_61,
$json_subdata_62,
$json_subdata_63,
$json_subdata_64,
$json_subdata_65,
$json_subdata_66,
$json_subdata_67,
$json_subdata_68,
$json_subdata_69,
$json_subdata_70,
$json_subdata_71,
$json_subdata_72,
$json_subdata_73,
$json_subdata_74,
$json_subdata_75,
$json_subdata_76,
$json_subdata_77,
$json_subdata_78,
$json_subdata_79,
$json_subdata_80
];

$json_draw_4['id'] = 'D31';
$json_draw_4['value'] = $desired_coal_hhv_value. ' '. $desired_coal_hhv_unit; 
$json_draw_5['id'] = 'D53';
$json_draw_5['value'] = $gross_output_value. ' '. $gross_output_unit; 
$json_draw_17['id'] = 'D41';
$json_draw_17['value'] = $main_steam_inlet_temperature_value. ' '. $main_steam_inlet_temperature_unit; 
$json_draw_18['id'] = 'D42';
$json_draw_18['value'] = $main_steam_inlet_pressure_value. ' '. $main_steam_inlet_pressure_unit; 
$json_draw_19['id'] = 'D40';
$json_draw_19['value'] = $main_steam_inlet_flow_value. ' '. $main_steam_inlet_flow_unit; 
$json_draw_20['id'] = 'D44';
$json_draw_20['value'] = $hrh_steam_inlet_temperature_value. ' '. $hrh_steam_inlet_temperature_unit; 
$json_draw_21['id'] = 'D45';
$json_draw_21['value'] = $hrh_steam_inlet_pressure_value. ' '. $hrh_steam_inlet_pressure_unit; 
$json_draw_22['id'] = 'D43';
$json_draw_22['value'] = $hrh_steam_inlet_flow_value. ' '. $hrh_steam_inlet_flow_unit; 
$json_draw_23['id'] = 'D51';
$json_draw_23['value'] = $lp_st_steam_inlet_temperature_value. ' '. $lp_st_steam_inlet_temperature_unit; 
$json_draw_24['id'] = 'D52';
$json_draw_24['value'] = $lp_st_steam_inlet_pressure_value. ' '. $lp_st_steam_inlet_pressure_unit; 
$json_draw_25['id'] = 'D50';
$json_draw_25['value'] = $lp_st_steam_inlet_flow_value. ' '. $lp_st_steam_inlet_flow_unit; 
$json_draw_26['id'] = 'D20';
$json_draw_26['value'] = $cond_cw_flow_value. ' '. $cond_cw_flow_unit; 
$json_draw_36['id'] = 'D13';
$json_draw_36['value'] = $fwh1_terminal_temperature_difference_value. ' '. $fwh1_terminal_temperature_difference_unit; 
$json_draw_37['id'] = 'D11';
$json_draw_37['value'] = $fwh2_terminal_temperature_difference_value. ' '. $fwh2_terminal_temperature_difference_unit; 
$json_draw_38['id'] = 'D09';
$json_draw_38['value'] = $fwh3_terminal_temperature_difference_value. ' '. $fwh3_terminal_temperature_difference_unit; 
$json_draw_39['id'] = 'D07';
$json_draw_39['value'] = $fwh5_terminal_temperature_difference_value. ' '. $fwh5_terminal_temperature_difference_unit; 
$json_draw_40['id'] = 'D05';
$json_draw_40['value'] = $fwh6_terminal_temperature_difference_value. ' '. $fwh6_terminal_temperature_difference_unit; 
$json_draw_41['id'] = 'D03';
$json_draw_41['value'] = $fwh7_terminal_temperature_difference_value. ' '. $fwh7_terminal_temperature_difference_unit; 
$json_draw_42['id'] = 'D01';
$json_draw_42['value'] = $fwh8_terminal_temperature_difference_value. ' '. $fwh8_terminal_temperature_difference_unit; 
$json_draw_43['id'] = 'D14';
$json_draw_43['value'] = $fwh1_drain_cooler_approach_value. ' '. $fwh1_drain_cooler_approach_unit; 
$json_draw_44['id'] = 'D12';
$json_draw_44['value'] = $fwh2_drain_cooler_approach_value. ' '. $fwh2_drain_cooler_approach_unit; 
$json_draw_45['id'] = 'D10';
$json_draw_45['value'] = $fwh3_drain_cooler_approach_value. ' '. $fwh3_drain_cooler_approach_unit; 
$json_draw_46['id'] = 'D08';
$json_draw_46['value'] = $fwh5_drain_cooler_approach_value. ' '. $fwh5_drain_cooler_approach_unit; 
$json_draw_47['id'] = 'D06';
$json_draw_47['value'] = $fwh6_drain_cooler_approach_value. ' '. $fwh6_drain_cooler_approach_unit; 
$json_draw_48['id'] = 'D04';
$json_draw_48['value'] = $fwh7_drain_cooler_approach_value. ' '. $fwh7_drain_cooler_approach_unit; 
$json_draw_49['id'] = 'D02';
$json_draw_49['value'] = $fwh8_drain_cooler_approach_value. ' '. $fwh8_drain_cooler_approach_unit; 
$json_draw_50['id'] = 'D18';
$json_draw_50['value'] = $cond_main_steam_inlet_pressure_value. ' '. $cond_main_steam_inlet_pressure_unit; 
$json_draw_52['id'] = 'D19';
$json_draw_52['value'] = $cond_cooling_water_inlet_temperature_value. ' '. $cond_cooling_water_inlet_temperature_unit; 
$json_draw_53['id'] = 'D21';
$json_draw_53['value'] = $cond_cooling_water_exit_temperature_value. ' '. $cond_cooling_water_exit_temperature_unit; 
$json_draw_54['id'] = 'D20';
$json_draw_54['value'] = $cond_cooling_water_inlet_flow_value. ' '. $cond_cooling_water_inlet_flow_unit; 
$json_draw_58['id'] = 'D30';
$json_draw_58['value'] = $coal_flow_value. ' '. $coal_flow_unit; 
$json_draw_59['id'] = 'D38';
$json_draw_59['value'] = $stack_temperature_value. ' '. $stack_temperature_unit; 
$json_draw_63['id'] = 'D32';
$json_draw_63['value'] = $aph_flue_gas_inlet_temperature_value. ' '. $aph_flue_gas_inlet_temperature_unit; 
$json_draw_64['id'] = 'D34';
$json_draw_64['value'] = $sa_outlet_temp_value. ' '. $sa_outlet_temp_unit; 
$json_draw_65['id'] = 'D36';
$json_draw_65['value'] = $pa_outlet_temp_value. ' '. $pa_outlet_temp_unit; 
$json_draw_66['id'] = 'D33';
$json_draw_66['value'] = $flue_gas_outlet_temp_value. ' '. $flue_gas_outlet_temp_unit; 
$json_draw_67['id'] = 'D35';
$json_draw_67['value'] = $sa_inlet_temp_value. ' '. $sa_inlet_temp_unit; 
$json_draw_68['id'] = 'D37';
$json_draw_68['value'] = $pa_inlet_temp_value. ' '. $pa_inlet_temp_unit; 
$json_draw_69['id'] = 'D46';
$json_draw_69['value'] = $crh_inlet_mass_flow_value. ' '. $crh_inlet_mass_flow_unit; 
$json_draw_70['id'] = 'D47';
$json_draw_70['value'] = $crh_inlet_temp_value. ' '. $crh_inlet_temp_unit; 
$json_draw_71['id'] = 'D48';
$json_draw_71['value'] = $crh_inlet_pressure_value. ' '. $crh_inlet_pressure_unit; 
$json_draw_72['id'] = 'D55';
$json_draw_72['value'] = $lp_steam_efficiency_value. ' '. $lp_steam_efficiency_unit; 
$json_draw_73['id'] = 'D56';
$json_draw_73['value'] = $ip_steam_efficiency_value. ' '. $ip_steam_efficiency_unit; 
$json_draw_74['id'] = 'D57';
$json_draw_74['value'] = $hp_steam_efficiency_value. ' '. $hp_steam_efficiency_unit; 
$json_draw_75['id'] = 'D58';
$json_draw_75['value'] = $ambient_temperature_value. ' '. $ambient_temperature_unit; 
$json_draw_76['id'] = 'D59';
$json_draw_76['value'] = $relative_humidity_value. ' '. $relative_humidity_unit; 
$json_draw_77['id'] = 'D60';
$json_draw_77['value'] = $ambient_pressure_value. ' '. $ambient_pressure_unit; 
$json_draw_78['id'] = 'D15';
$json_draw_78['value'] = $feedwater_mass_flow_value. ' '. $feedwater_mass_flow_unit; 
$json_draw_79['id'] = 'D16';
$json_draw_79['value'] = $feedwater_temperature_value. ' '. $feedwater_temperature_unit; 
$json_draw_80['id'] = 'D17';
$json_draw_80['value'] = $feedwater_pressure_value. ' '. $feedwater_pressure_unit; 

$json_draw_81['id'] = 'D61';
$json_draw_81['value'] = $data_input['ambient_temp']. ' '. $ambient_temperature_unit; 

$json_draw_82['id'] = 'D62';
$json_draw_82['value'] = $data_input['relative_humidity']. ' '. $relative_humidity_unit; 

$json_draw_83['id'] = 'D63';
$json_draw_83['value'] = $data_input['ambient_press']. ' '. $ambient_pressure_unit; 

$json_data['drawing'] = [
	$json_draw_4,
	$json_draw_5,
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
	$json_draw_47,
	$json_draw_48,
	$json_draw_49,
	$json_draw_50,
	$json_draw_52,
	$json_draw_53,
	$json_draw_54,
	$json_draw_58,
	$json_draw_59,
	$json_draw_63,
	$json_draw_64,
	$json_draw_65,
	$json_draw_66,
	$json_draw_67,
	$json_draw_68,
	$json_draw_69,
	$json_draw_70,
	$json_draw_71,
	$json_draw_72,
	$json_draw_73,
	$json_draw_74,
	$json_draw_75,
	$json_draw_76,
	$json_draw_77,
	$json_draw_78,
	$json_draw_79,
	$json_draw_80,
	$json_draw_81,
	$json_draw_82,
	$json_draw_83
];

//5. serve output value output value will skip after error.
echo json_encode($json_data);
} catch (Exception $e){
	header("HTTP/1.1 500 Internal Server Error");
	echo $e->getMessage();
	die();
}
?>