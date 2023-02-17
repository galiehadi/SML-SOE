Set objExcel_tja_master16359951300423    = CreateObject("Excel.Application")
Set objWorkbook_tja_master16359951300423 = objExcel_tja_master16359951300423.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\tja_master16359951300423.xlsm")
objExcel_tja_master16359951300423.Range("input!E1").Value = "300"
objExcel_tja_master16359951300423.Range("input!N3").Value = "0.9"
objExcel_tja_master16359951300423.Range("input!N4").Value = "17584.56"
objExcel_tja_master16359951300423.Range("input!N5").Value = "33"
objExcel_tja_master16359951300423.Range("input!N12").Value = "29"
objExcel_tja_master16359951300423.Range("input!N13").Value = "70"
objExcel_tja_master16359951300423.Range("input!N14").Value = "1.0124"
objExcel_tja_master16359951300423.Range("input!N11").Value = "25"
objExcel_tja_master16359951300423.Range("input!N26").Value = "28"
objExcel_tja_master16359951300423.Range("input!N27").Value = "10"
objWorkbook_tja_master16359951300423.Save
objExcel_tja_master16359951300423.Application.DisplayAlerts = False
objExcel_tja_master16359951300423.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\tja_master16359951300423.xlsm'!GetData"
objExcel_tja_master16359951300423.Application.DisplayAlerts = False
objWorkbook_tja_master16359951300423.Save
objWorkbook_tja_master16359951300423.Close False 
objExcel_tja_master16359951300423.Application.Quit
Set objExcel_tja_master16359951300423 = Nothing 
Set objWorkbook_tja_master16359951300423 = Nothing
