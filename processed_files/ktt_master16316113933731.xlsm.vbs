Set objExcel_ktt_master16316113933731    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316113933731 = objExcel_ktt_master16316113933731.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316113933731.xlsm")
objExcel_ktt_master16316113933731.Range("input!C1").Value = "55"
objExcel_ktt_master16316113933731.Range("input!J4").Value = "55000"
objExcel_ktt_master16316113933731.Range("input!J3").Value = "0.8"
objExcel_ktt_master16316113933731.Range("input!J5").Value = "14653.8"
objExcel_ktt_master16316113933731.Range("input!J6").Value = "25"
objExcel_ktt_master16316113933731.Range("input!J7").Value = "25"
objExcel_ktt_master16316113933731.Range("input!J8").Value = "55"
objExcel_ktt_master16316113933731.Range("input!J9").Value = "1.0122"
objExcel_ktt_master16316113933731.Range("input!J10").Value = "15"
objExcel_ktt_master16316113933731.Range("input!J11").Value = "27"
objExcel_ktt_master16316113933731.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316113933731.Save
objExcel_ktt_master16316113933731.Application.DisplayAlerts = False
objExcel_ktt_master16316113933731.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316113933731.xlsm'!GetData"
objExcel_ktt_master16316113933731.Application.DisplayAlerts = False
objWorkbook_ktt_master16316113933731.Save
objWorkbook_ktt_master16316113933731.Close False 
objExcel_ktt_master16316113933731.Application.Quit
Set objExcel_ktt_master16316113933731 = Nothing 
Set objWorkbook_ktt_master16316113933731 = Nothing
