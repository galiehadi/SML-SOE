Set objExcel_ktt_master16316193046474    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316193046474 = objExcel_ktt_master16316193046474.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316193046474.xlsm")
objExcel_ktt_master16316193046474.Range("input!C1").Value = "92"
objExcel_ktt_master16316193046474.Range("input!J4").Value = "92000"
objExcel_ktt_master16316193046474.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316193046474.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316193046474.Range("input!J6").Value = "34"
objExcel_ktt_master16316193046474.Range("input!J7").Value = "27"
objExcel_ktt_master16316193046474.Range("input!J8").Value = "63"
objExcel_ktt_master16316193046474.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316193046474.Range("input!J10").Value = "22"
objExcel_ktt_master16316193046474.Range("input!J11").Value = "27"
objExcel_ktt_master16316193046474.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316193046474.Save
objExcel_ktt_master16316193046474.Application.DisplayAlerts = False
objExcel_ktt_master16316193046474.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316193046474.xlsm'!GetData"
objExcel_ktt_master16316193046474.Application.DisplayAlerts = False
objWorkbook_ktt_master16316193046474.Save
objWorkbook_ktt_master16316193046474.Close False 
objExcel_ktt_master16316193046474.Application.Quit
Set objExcel_ktt_master16316193046474 = Nothing 
Set objWorkbook_ktt_master16316193046474 = Nothing
