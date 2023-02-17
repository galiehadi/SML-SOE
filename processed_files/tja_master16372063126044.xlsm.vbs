Set objExcel_tja_master16372063126044    = CreateObject("Excel.Application")
Set objWorkbook_tja_master16372063126044 = objExcel_tja_master16372063126044.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\tja_master16372063126044.xlsm")
objExcel_tja_master16372063126044.Range("input!E1").Value = ""
objExcel_tja_master16372063126044.Range("input!N3").Value = ""
objExcel_tja_master16372063126044.Range("input!N4").Value = "0"
objExcel_tja_master16372063126044.Range("input!N5").Value = ""
objExcel_tja_master16372063126044.Range("input!N12").Value = ""
objExcel_tja_master16372063126044.Range("input!N13").Value = ""
objExcel_tja_master16372063126044.Range("input!N14").Value = ""
objExcel_tja_master16372063126044.Range("input!N11").Value = ""
objExcel_tja_master16372063126044.Range("input!N26").Value = ""
objExcel_tja_master16372063126044.Range("input!N27").Value = ""
objWorkbook_tja_master16372063126044.Save
objExcel_tja_master16372063126044.Application.DisplayAlerts = False
objExcel_tja_master16372063126044.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\tja_master16372063126044.xlsm'!GetData"
objExcel_tja_master16372063126044.Application.DisplayAlerts = False
objWorkbook_tja_master16372063126044.Save
objWorkbook_tja_master16372063126044.Close False 
objExcel_tja_master16372063126044.Application.Quit
Set objExcel_tja_master16372063126044 = Nothing 
Set objWorkbook_tja_master16372063126044 = Nothing
