Set objExcel_ktt_master16316142043199    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316142043199 = objExcel_ktt_master16316142043199.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316142043199.xlsm")
objExcel_ktt_master16316142043199.Range("input!C1").Value = "75"
objExcel_ktt_master16316142043199.Range("input!J4").Value = "75000"
objExcel_ktt_master16316142043199.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316142043199.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316142043199.Range("input!J6").Value = "34"
objExcel_ktt_master16316142043199.Range("input!J7").Value = "27"
objExcel_ktt_master16316142043199.Range("input!J8").Value = "63"
objExcel_ktt_master16316142043199.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316142043199.Range("input!J10").Value = "22"
objExcel_ktt_master16316142043199.Range("input!J11").Value = "27"
objExcel_ktt_master16316142043199.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316142043199.Save
objExcel_ktt_master16316142043199.Application.DisplayAlerts = False
objExcel_ktt_master16316142043199.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316142043199.xlsm'!GetData"
objExcel_ktt_master16316142043199.Application.DisplayAlerts = False
objWorkbook_ktt_master16316142043199.Save
objWorkbook_ktt_master16316142043199.Close False 
objExcel_ktt_master16316142043199.Application.Quit
Set objExcel_ktt_master16316142043199 = Nothing 
Set objWorkbook_ktt_master16316142043199 = Nothing
