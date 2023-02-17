Set objExcel_amg_master16697112858057    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16697112858057 = objExcel_amg_master16697112858057.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16697112858057.xlsm")
objExcel_amg_master16697112858057.Range("input!C1").Value = "12.5"
objExcel_amg_master16697112858057.Range("input!J4").Value = "12500"
objExcel_amg_master16697112858057.Range("input!J3").Value = "0.8"
objExcel_amg_master16697112858057.Range("input!J5").Value = "14653.8"
objExcel_amg_master16697112858057.Range("input!J6").Value = "25"
objExcel_amg_master16697112858057.Range("input!J7").Value = "25"
objExcel_amg_master16697112858057.Range("input!J8").Value = "55"
objExcel_amg_master16697112858057.Range("input!J9").Value = "1.0122"
objExcel_amg_master16697112858057.Range("input!J10").Value = "15"
objExcel_amg_master16697112858057.Range("input!J11").Value = "27"
objExcel_amg_master16697112858057.Range("input!J12").Value = "8"
objWorkbook_amg_master16697112858057.Save
objExcel_amg_master16697112858057.Application.DisplayAlerts = False
objExcel_amg_master16697112858057.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16697112858057.xlsm'!GetData"
objExcel_amg_master16697112858057.Application.DisplayAlerts = False
objWorkbook_amg_master16697112858057.Save
objWorkbook_amg_master16697112858057.Close False 
objExcel_amg_master16697112858057.Application.Quit
Set objExcel_amg_master16697112858057 = Nothing 
Set objWorkbook_amg_master16697112858057 = Nothing
