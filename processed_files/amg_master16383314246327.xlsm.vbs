Set objExcel_amg_master16383314246327    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16383314246327 = objExcel_amg_master16383314246327.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16383314246327.xlsm")
objExcel_amg_master16383314246327.Range("input!C1").Value = "24.5"
objExcel_amg_master16383314246327.Range("input!J4").Value = "24500"
objExcel_amg_master16383314246327.Range("input!J3").Value = "0.9"
objExcel_amg_master16383314246327.Range("input!J5").Value = "16328.52"
objExcel_amg_master16383314246327.Range("input!J6").Value = "33"
objExcel_amg_master16383314246327.Range("input!J7").Value = "29"
objExcel_amg_master16383314246327.Range("input!J8").Value = "65"
objExcel_amg_master16383314246327.Range("input!J9").Value = "1.0124"
objExcel_amg_master16383314246327.Range("input!J10").Value = "20"
objExcel_amg_master16383314246327.Range("input!J11").Value = "29"
objExcel_amg_master16383314246327.Range("input!J12").Value = "10"
objWorkbook_amg_master16383314246327.Save
objExcel_amg_master16383314246327.Application.DisplayAlerts = False
objExcel_amg_master16383314246327.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16383314246327.xlsm'!GetData"
objExcel_amg_master16383314246327.Application.DisplayAlerts = False
objWorkbook_amg_master16383314246327.Save
objWorkbook_amg_master16383314246327.Close False 
objExcel_amg_master16383314246327.Application.Quit
Set objExcel_amg_master16383314246327 = Nothing 
Set objWorkbook_amg_master16383314246327 = Nothing
