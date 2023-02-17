Set objExcel_ktt_master16316286967448    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316286967448 = objExcel_ktt_master16316286967448.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316286967448.xlsm")
objExcel_ktt_master16316286967448.Range("input!C1").Value = "60"
objExcel_ktt_master16316286967448.Range("input!J4").Value = "60000"
objExcel_ktt_master16316286967448.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316286967448.Range("input!J5").Value = "16747.2"
objExcel_ktt_master16316286967448.Range("input!J6").Value = "33"
objExcel_ktt_master16316286967448.Range("input!J7").Value = "27"
objExcel_ktt_master16316286967448.Range("input!J8").Value = "63"
objExcel_ktt_master16316286967448.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316286967448.Range("input!J10").Value = "20"
objExcel_ktt_master16316286967448.Range("input!J11").Value = "27"
objExcel_ktt_master16316286967448.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316286967448.Save
objExcel_ktt_master16316286967448.Application.DisplayAlerts = False
objExcel_ktt_master16316286967448.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316286967448.xlsm'!GetData"
objExcel_ktt_master16316286967448.Application.DisplayAlerts = False
objWorkbook_ktt_master16316286967448.Save
objWorkbook_ktt_master16316286967448.Close False 
objExcel_ktt_master16316286967448.Application.Quit
Set objExcel_ktt_master16316286967448 = Nothing 
Set objWorkbook_ktt_master16316286967448 = Nothing
