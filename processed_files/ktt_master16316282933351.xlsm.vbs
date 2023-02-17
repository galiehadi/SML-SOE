Set objExcel_ktt_master16316282933351    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316282933351 = objExcel_ktt_master16316282933351.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316282933351.xlsm")
objExcel_ktt_master16316282933351.Range("input!C1").Value = "88"
objExcel_ktt_master16316282933351.Range("input!J4").Value = "88000"
objExcel_ktt_master16316282933351.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316282933351.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316282933351.Range("input!J6").Value = "34"
objExcel_ktt_master16316282933351.Range("input!J7").Value = "27"
objExcel_ktt_master16316282933351.Range("input!J8").Value = "63"
objExcel_ktt_master16316282933351.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316282933351.Range("input!J10").Value = "22"
objExcel_ktt_master16316282933351.Range("input!J11").Value = "27"
objExcel_ktt_master16316282933351.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316282933351.Save
objExcel_ktt_master16316282933351.Application.DisplayAlerts = False
objExcel_ktt_master16316282933351.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316282933351.xlsm'!GetData"
objExcel_ktt_master16316282933351.Application.DisplayAlerts = False
objWorkbook_ktt_master16316282933351.Save
objWorkbook_ktt_master16316282933351.Close False 
objExcel_ktt_master16316282933351.Application.Quit
Set objExcel_ktt_master16316282933351 = Nothing 
Set objWorkbook_ktt_master16316282933351 = Nothing
