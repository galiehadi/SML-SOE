Set objExcel_ktt_master16365101686752    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master16365101686752 = objExcel_ktt_master16365101686752.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master16365101686752.xlsm")
objExcel_ktt_master16365101686752.Range("input!C1").Value = "90"
objExcel_ktt_master16365101686752.Range("input!J4").Value = "90000"
objExcel_ktt_master16365101686752.Range("input!J3").Value = "0.85"
objExcel_ktt_master16365101686752.Range("input!J5").Value = "15909.84"
objExcel_ktt_master16365101686752.Range("input!J6").Value = "34"
objExcel_ktt_master16365101686752.Range("input!J7").Value = "27"
objExcel_ktt_master16365101686752.Range("input!J8").Value = "80"
objExcel_ktt_master16365101686752.Range("input!J9").Value = "1.0123"
objExcel_ktt_master16365101686752.Range("input!J10").Value = "22"
objExcel_ktt_master16365101686752.Range("input!J11").Value = "30"
objExcel_ktt_master16365101686752.Range("input!J12").Value = "8"
objWorkbook_ktt_master16365101686752.Save
objExcel_ktt_master16365101686752.Application.DisplayAlerts = False
objExcel_ktt_master16365101686752.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master16365101686752.xlsm'!GetData"
objExcel_ktt_master16365101686752.Application.DisplayAlerts = False
objWorkbook_ktt_master16365101686752.Save
objWorkbook_ktt_master16365101686752.Close False 
objExcel_ktt_master16365101686752.Application.Quit
Set objExcel_ktt_master16365101686752 = Nothing 
Set objWorkbook_ktt_master16365101686752 = Nothing
