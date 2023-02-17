Set objExcel_ktt_master16376763743229    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16376763743229 = objExcel_ktt_master16376763743229.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16376763743229.xlsm")
objExcel_ktt_master16376763743229.Range("input!C1").Value = "100"
objExcel_ktt_master16376763743229.Range("input!J4").Value = "100000"
objExcel_ktt_master16376763743229.Range("input!J3").Value = "0.85"
objExcel_ktt_master16376763743229.Range("input!J5").Value = "16747.2"
objExcel_ktt_master16376763743229.Range("input!J6").Value = "33"
objExcel_ktt_master16376763743229.Range("input!J7").Value = "27"
objExcel_ktt_master16376763743229.Range("input!J8").Value = "63"
objExcel_ktt_master16376763743229.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16376763743229.Range("input!J10").Value = "20"
objExcel_ktt_master16376763743229.Range("input!J11").Value = "27"
objExcel_ktt_master16376763743229.Range("input!J12").Value = "8"
objWorkbook_ktt_master16376763743229.Save
objExcel_ktt_master16376763743229.Application.DisplayAlerts = False
objExcel_ktt_master16376763743229.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16376763743229.xlsm'!GetData"
objExcel_ktt_master16376763743229.Application.DisplayAlerts = False
objWorkbook_ktt_master16376763743229.Save
objWorkbook_ktt_master16376763743229.Close False 
objExcel_ktt_master16376763743229.Application.Quit
Set objExcel_ktt_master16376763743229 = Nothing 
Set objWorkbook_ktt_master16376763743229 = Nothing
