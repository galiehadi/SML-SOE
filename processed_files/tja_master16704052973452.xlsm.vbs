Set objExcel_tja_master16704052973452    = CreateObject("Excel.Application")
Set objWorkbook_tja_master16704052973452 = objExcel_tja_master16704052973452.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\tja_master16704052973452.xlsm")
objExcel_tja_master16704052973452.Range("input!E1").Value = "300"
objExcel_tja_master16704052973452.Range("input!N3").Value = "0.9"
objExcel_tja_master16704052973452.Range("input!N4").Value = "17584.56"
objExcel_tja_master16704052973452.Range("input!N5").Value = "33"
objExcel_tja_master16704052973452.Range("input!N12").Value = "29"
objExcel_tja_master16704052973452.Range("input!N13").Value = "70"
objExcel_tja_master16704052973452.Range("input!N14").Value = "1.0124"
objExcel_tja_master16704052973452.Range("input!N11").Value = "25"
objExcel_tja_master16704052973452.Range("input!N26").Value = "28"
objExcel_tja_master16704052973452.Range("input!N27").Value = "10"
objWorkbook_tja_master16704052973452.Save
objExcel_tja_master16704052973452.Application.DisplayAlerts = False
objExcel_tja_master16704052973452.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\tja_master16704052973452.xlsm'!GetData"
objExcel_tja_master16704052973452.Application.DisplayAlerts = False
objWorkbook_tja_master16704052973452.Save
objWorkbook_tja_master16704052973452.Close False 
objExcel_tja_master16704052973452.Application.Quit
Set objExcel_tja_master16704052973452 = Nothing 
Set objWorkbook_tja_master16704052973452 = Nothing
