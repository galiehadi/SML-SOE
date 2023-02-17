Set objExcel_ktt_master16350416447251    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16350416447251 = objExcel_ktt_master16350416447251.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16350416447251.xlsm")
objExcel_ktt_master16350416447251.Range("input!C1").Value = "95"
objExcel_ktt_master16350416447251.Range("input!J4").Value = "95000"
objExcel_ktt_master16350416447251.Range("input!J3").Value = "0.85"
objExcel_ktt_master16350416447251.Range("input!J5").Value = "16747.2"
objExcel_ktt_master16350416447251.Range("input!J6").Value = "33"
objExcel_ktt_master16350416447251.Range("input!J7").Value = "27"
objExcel_ktt_master16350416447251.Range("input!J8").Value = "63"
objExcel_ktt_master16350416447251.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16350416447251.Range("input!J10").Value = "20"
objExcel_ktt_master16350416447251.Range("input!J11").Value = "27"
objExcel_ktt_master16350416447251.Range("input!J12").Value = "8"
objWorkbook_ktt_master16350416447251.Save
objExcel_ktt_master16350416447251.Application.DisplayAlerts = False
objExcel_ktt_master16350416447251.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16350416447251.xlsm'!GetData"
objExcel_ktt_master16350416447251.Application.DisplayAlerts = False
objWorkbook_ktt_master16350416447251.Save
objWorkbook_ktt_master16350416447251.Close False 
objExcel_ktt_master16350416447251.Application.Quit
Set objExcel_ktt_master16350416447251 = Nothing 
Set objWorkbook_ktt_master16350416447251 = Nothing
