Set objExcel_amg_master16389434838886    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16389434838886 = objExcel_amg_master16389434838886.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16389434838886.xlsm")
objExcel_amg_master16389434838886.Range("input!C1").Value = "15"
objExcel_amg_master16389434838886.Range("input!J4").Value = "15000"
objExcel_amg_master16389434838886.Range("input!J3").Value = "0.88"
objExcel_amg_master16389434838886.Range("input!J5").Value = "16747.2"
objExcel_amg_master16389434838886.Range("input!J6").Value = "33"
objExcel_amg_master16389434838886.Range("input!J7").Value = "28"
objExcel_amg_master16389434838886.Range("input!J8").Value = "67"
objExcel_amg_master16389434838886.Range("input!J9").Value = "1.0124"
objExcel_amg_master16389434838886.Range("input!J10").Value = "20"
objExcel_amg_master16389434838886.Range("input!J11").Value = "30"
objExcel_amg_master16389434838886.Range("input!J12").Value = "9"
objWorkbook_amg_master16389434838886.Save
objExcel_amg_master16389434838886.Application.DisplayAlerts = False
objExcel_amg_master16389434838886.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16389434838886.xlsm'!GetData"
objExcel_amg_master16389434838886.Application.DisplayAlerts = False
objWorkbook_amg_master16389434838886.Save
objWorkbook_amg_master16389434838886.Close False 
objExcel_amg_master16389434838886.Application.Quit
Set objExcel_amg_master16389434838886 = Nothing 
Set objWorkbook_amg_master16389434838886 = Nothing
