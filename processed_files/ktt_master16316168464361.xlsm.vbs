Set objExcel_ktt_master16316168464361    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16316168464361 = objExcel_ktt_master16316168464361.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16316168464361.xlsm")
objExcel_ktt_master16316168464361.Range("input!C1").Value = "95"
objExcel_ktt_master16316168464361.Range("input!J4").Value = "95000"
objExcel_ktt_master16316168464361.Range("input!J3").Value = "0.85"
objExcel_ktt_master16316168464361.Range("input!J5").Value = "16642.53"
objExcel_ktt_master16316168464361.Range("input!J6").Value = "34"
objExcel_ktt_master16316168464361.Range("input!J7").Value = "27"
objExcel_ktt_master16316168464361.Range("input!J8").Value = "63"
objExcel_ktt_master16316168464361.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16316168464361.Range("input!J10").Value = "22"
objExcel_ktt_master16316168464361.Range("input!J11").Value = "27"
objExcel_ktt_master16316168464361.Range("input!J12").Value = "8"
objWorkbook_ktt_master16316168464361.Save
objExcel_ktt_master16316168464361.Application.DisplayAlerts = False
objExcel_ktt_master16316168464361.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16316168464361.xlsm'!GetData"
objExcel_ktt_master16316168464361.Application.DisplayAlerts = False
objWorkbook_ktt_master16316168464361.Save
objWorkbook_ktt_master16316168464361.Close False 
objExcel_ktt_master16316168464361.Application.Quit
Set objExcel_ktt_master16316168464361 = Nothing 
Set objWorkbook_ktt_master16316168464361 = Nothing
