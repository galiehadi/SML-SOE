Set objExcel_ktt_master16359973357116    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16359973357116 = objExcel_ktt_master16359973357116.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16359973357116.xlsm")
objExcel_ktt_master16359973357116.Range("input!C1").Value = "100"
objExcel_ktt_master16359973357116.Range("input!J4").Value = "100000"
objExcel_ktt_master16359973357116.Range("input!J3").Value = "0.85"
objExcel_ktt_master16359973357116.Range("input!J5").Value = "15909.84"
objExcel_ktt_master16359973357116.Range("input!J6").Value = "34"
objExcel_ktt_master16359973357116.Range("input!J7").Value = "30"
objExcel_ktt_master16359973357116.Range("input!J8").Value = "70"
objExcel_ktt_master16359973357116.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16359973357116.Range("input!J10").Value = "22"
objExcel_ktt_master16359973357116.Range("input!J11").Value = "32"
objExcel_ktt_master16359973357116.Range("input!J12").Value = "8"
objWorkbook_ktt_master16359973357116.Save
objExcel_ktt_master16359973357116.Application.DisplayAlerts = False
objExcel_ktt_master16359973357116.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16359973357116.xlsm'!GetData"
objExcel_ktt_master16359973357116.Application.DisplayAlerts = False
objWorkbook_ktt_master16359973357116.Save
objWorkbook_ktt_master16359973357116.Close False 
objExcel_ktt_master16359973357116.Application.Quit
Set objExcel_ktt_master16359973357116 = Nothing 
Set objWorkbook_ktt_master16359973357116 = Nothing
