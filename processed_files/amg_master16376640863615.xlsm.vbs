Set objExcel_amg_master16376640863615    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16376640863615 = objExcel_amg_master16376640863615.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16376640863615.xlsm")
objExcel_amg_master16376640863615.Range("input!C1").Value = "20"
objExcel_amg_master16376640863615.Range("input!J4").Value = "20000"
objExcel_amg_master16376640863615.Range("input!J3").Value = "0.9"
objExcel_amg_master16376640863615.Range("input!J5").Value = "17584.56"
objExcel_amg_master16376640863615.Range("input!J6").Value = "33"
objExcel_amg_master16376640863615.Range("input!J7").Value = "29"
objExcel_amg_master16376640863615.Range("input!J8").Value = "70"
objExcel_amg_master16376640863615.Range("input!J9").Value = "1.0124"
objExcel_amg_master16376640863615.Range("input!J10").Value = "25"
objExcel_amg_master16376640863615.Range("input!J11").Value = "28"
objExcel_amg_master16376640863615.Range("input!J12").Value = "10"
objWorkbook_amg_master16376640863615.Save
objExcel_amg_master16376640863615.Application.DisplayAlerts = False
objExcel_amg_master16376640863615.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16376640863615.xlsm'!GetData"
objExcel_amg_master16376640863615.Application.DisplayAlerts = False
objWorkbook_amg_master16376640863615.Save
objWorkbook_amg_master16376640863615.Close False 
objExcel_amg_master16376640863615.Application.Quit
Set objExcel_amg_master16376640863615 = Nothing 
Set objWorkbook_amg_master16376640863615 = Nothing
