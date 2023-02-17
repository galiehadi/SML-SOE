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
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(488,5).Value = "'.$data_input['generator_power_factor'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(238,5).Value = "'.$data_input['desired_coal_hhv'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(230,5).Value = "'.$data_input['total_moisture'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(429,5).Value = "'.$data_input['main_steam_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(256,5).Value = "'.$data_input['main_steam_mass_flow'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(29,5).Value = "'.$data_input['desup_2_outlet_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(449,5).Value = "'.$data_input['platen_sh_temp_out'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(445,5).Value = "'.$data_input['panel_sh_temp_out'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(27,5).Value = "'.$data_input['desup_1_out_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(433,5).Value = "'.$data_input['ltsh_out_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(37,5).Value = "'.$data_input['econ_out_temp'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(509,5).Value = "'.$data_input['main_steam_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(484,5).Value = "'.$data_input['cold_reheat_press'].'"'."\r\n"; 
$vbs_content = $vbs_content. 'objExcel_'.$fileRep.'.Cells(485,5).Value = "'.$data_input['hrh_press'].'"'."\r\n"; 
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
//defaine output end

$servername = "localhost";
$username = "root";
$password = "";
$dbname = "sampling";

// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);
// Check connection
if ($conn->connect_error) {
  die("Connection failed: " . $conn->connect_error);
}

$sql = "INSERT INTO elink_sampling (
generator_power_factor,
desired_coal_hhv,
total_moisture,
main_steam_temp,
main_steam_mass_flow,
desup_2_outlet_temp,
platen_sh_temp_out,
panel_sh_temp_out,
desup_1_out_temp,
ltsh_out_temp,
econ_out_temp,
main_steam_press,
cold_reheat_press,
hrh_press,
expected_power_output,
auxiliary_power_consumption,
gross_output,
boiler_efficiency,
net_output,
net_heatrate,
gross_heatrate,
plant_efficiency,
ssh_steam_outlet_temperature,
ssh_steam_outlet_pressure,
ssh_steam_outlet_flow,
frh_steam_outlet_temperature,
frh_steam_outlet_pressure,
frh_steam_outlet_flow,
main_steam_inlet_temperature,
main_steam_inlet_pressure,
main_steam_inlet_flow,
hrh_steam_inlet_temperature,
hrh_steam_inlet_pressure,
hrh_steam_inlet_flow,
lp_st_steam_inlet_temperature,
lp_st_steam_inlet_pressure,
lp_st_steam_inlet_flow,
cond_cw_flow,
hp_dry_step_eff1,
hp_dry_step_eff2,
ip_dry_step_eff1,
ip_dry_step_eff2,
lp_dry_step_eff1,
lp_dry_step_eff2,
lp_dry_step_eff3,
lp_dry_step_eff4,
lp_dry_step_eff5,
fwh1_terminal_temperature_difference,
fwh2_terminal_temperature_difference,
fwh3_terminal_temperature_difference,
fwh5_terminal_temperature_difference,
fwh6_terminal_temperature_difference,
fwh7_terminal_temperature_difference,
fwh8_terminal_temperature_difference,
fwh1_drain_cooler_approach,
fwh2_drain_cooler_approach,
fwh3_drain_cooler_approach,
fwh5_drain_cooler_approach,
fwh6_drain_cooler_approach,
fwh7_drain_cooler_approach,
fwh8_drain_cooler_approach,
cond_main_steam_inlet_pressure,
cond_calculated_effectiveness,
cond_cooling_water_inlet_temperature,
cond_cooling_water_exit_temperature,
cond_cooling_water_inlet_flow,
hrh_steam_outlet_temperature,
hrh_steam_outlet_pressure,
hrh_steam_outlet_flow,
coal_flow,
stack_temperature,
ambient_pressure,
ambient_rh,
cond_condensate_temperature,
aph_flue_gas_inlet_temperature,
sa_outlet_temp,
pa_outlet_temp
)
VALUES (
	".$spreadsheet->getActiveSheet()->getCell('E488')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E238')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E230')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E429')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E256')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E29')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E449')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E445')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E27')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E433')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E37')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E509')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E484')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E485')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E598')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E601')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E598')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1986')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E599')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E612')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E604')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E603')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1535')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1533')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1532')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1589')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1587')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1590')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2811')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2810')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2812')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2866')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2865')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2867')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2855')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2854')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2856')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E3692')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1737')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1764')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1791')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1818')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1845')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1872')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1899')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1926')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1958')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E917')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E926')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E935')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E971')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E962')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E953')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E944')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E918')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E927')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E936')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E972')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E963')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E954')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E945')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E3349')->getValue().",0,
	".$spreadsheet->getActiveSheet()->getCell('E3691')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E3273')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E3692')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2867')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2867')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2867')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E978')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2035')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E595')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E596')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E1664')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2170')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2305')->getValue().",
	".$spreadsheet->getActiveSheet()->getCell('E2170')->getValue()."
)";

if ($conn->query($sql) === TRUE) {
  echo "New record created successfully";
} else {
  echo "Error: " . $sql . "<br>" . $conn->error;
}

$conn->close();

$json_data = [];

//5. serve output value
echo json_encode($json_data);

?>