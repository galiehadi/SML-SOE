Set objExcel_ktt_master16316783302483    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316783302483 = objExcel_ktt_master16316783302483.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316783302483.xlsm")
objExcel_ktt_master16316783302483.Range("input!C1").Value = "75"
objExcel_ktt_master16316783302483.Range("input!J4").Value = "75000"
objExcel_ktt_master16316783302483.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316783302483.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316783302483.Range("input!J6").Value = "34"
objExcel_ktt_master16316783302483.Range("input!J7").Value = "27"
objExcel_ktt_master16316783302483.Range("input!J8").Value = "63"
objExcel_ktt_master16316783302483.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316783302483.Range("input!J10").Value = "22"
objExcel_ktt_master16316783302483.Range("input!J11").Value = "27"
objExcel_ktt_master16316783302483.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316783302483.Save
objExcel_ktt_master16316783302483.Application.DisplayAlerts = False
objExcel_ktt_master16316783302483.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316783302483.xlsm'!GetData"
objExcel_ktt_master16316783302483.Application.DisplayAlerts = False
objWorkbook_ktt_master16316783302483.Save
objWorkbook_ktt_master16316783302483.Close False 
objExcel_ktt_master16316783302483.Application.Quit
Set objExcel_ktt_master16316783302483 = Nothing 
Set objWorkbook_ktt_master16316783302483 = Nothing
