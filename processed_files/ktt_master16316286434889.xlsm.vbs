Set objExcel_ktt_master16316286434889    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316286434889 = objExcel_ktt_master16316286434889.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316286434889.xlsm")
objExcel_ktt_master16316286434889.Range("input!C1").Value = "56"
objExcel_ktt_master16316286434889.Range("input!J4").Value = "56000"
objExcel_ktt_master16316286434889.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316286434889.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316286434889.Range("input!J6").Value = "34"
objExcel_ktt_master16316286434889.Range("input!J7").Value = "27"
objExcel_ktt_master16316286434889.Range("input!J8").Value = "63"
objExcel_ktt_master16316286434889.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316286434889.Range("input!J10").Value = "22"
objExcel_ktt_master16316286434889.Range("input!J11").Value = "27"
objExcel_ktt_master16316286434889.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316286434889.Save
objExcel_ktt_master16316286434889.Application.DisplayAlerts = False
objExcel_ktt_master16316286434889.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316286434889.xlsm'!GetData"
objExcel_ktt_master16316286434889.Application.DisplayAlerts = False
objWorkbook_ktt_master16316286434889.Save
objWorkbook_ktt_master16316286434889.Close False 
objExcel_ktt_master16316286434889.Application.Quit
Set objExcel_ktt_master16316286434889 = Nothing 
Set objWorkbook_ktt_master16316286434889 = Nothing
