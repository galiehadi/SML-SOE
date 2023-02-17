Set objExcel_tja_master16337723277877    = CreateObject("Excel.Application")
Set objWorkbook_tja_master16337723277877 = objExcel_tja_master16337723277877.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\tja_master16337723277877.xlsm")
objExcel_tja_master16337723277877.Range("input!E1").Value = "300"
objExcel_tja_master16337723277877.Range("input!N3").Value = "0.9"
objExcel_tja_master16337723277877.Range("input!N4").Value = "17584.56"
objExcel_tja_master16337723277877.Range("input!N5").Value = "33"
objExcel_tja_master16337723277877.Range("input!N12").Value = "29"
objExcel_tja_master16337723277877.Range("input!N13").Value = "70"
objExcel_tja_master16337723277877.Range("input!N14").Value = "1.0124"
objExcel_tja_master16337723277877.Range("input!N11").Value = "25"
objExcel_tja_master16337723277877.Range("input!N26").Value = "28"
objExcel_tja_master16337723277877.Range("input!N27").Value = "10"
objWorkbook_tja_master16337723277877.Save
objExcel_tja_master16337723277877.Application.DisplayAlerts = False
objExcel_tja_master16337723277877.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\tja_master16337723277877.xlsm'!GetData"
objExcel_tja_master16337723277877.Application.DisplayAlerts = False
objWorkbook_tja_master16337723277877.Save
objWorkbook_tja_master16337723277877.Close False 
objExcel_tja_master16337723277877.Application.Quit
Set objExcel_tja_master16337723277877 = Nothing 
Set objWorkbook_tja_master16337723277877 = Nothing
