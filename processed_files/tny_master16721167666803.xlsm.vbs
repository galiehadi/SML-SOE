Set objExcel_tny_master16721167666803    = CreateObject("Excel.Application")
Set objWorkbook_tny_master16721167666803 = objExcel_tny_master16721167666803.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\tny_master16721167666803.xlsm")
objExcel_tny_master16721167666803.Range("input!E1").Value = "100"
objExcel_tny_master16721167666803.Range("input!M4").Value = "0.9"
objExcel_tny_master16721167666803.Range("input!M5").Value = "16747.2"
objExcel_tny_master16721167666803.Range("input!M6").Value = "33"
objExcel_tny_master16721167666803.Range("input!M10").Value = "28"
objExcel_tny_master16721167666803.Range("input!M11").Value = "30"
objExcel_tny_master16721167666803.Range("input!M12").Value = "70"
objExcel_tny_master16721167666803.Range("input!M13").Value = "1.0124"
objExcel_tny_master16721167666803.Range("input!M24").Value = "29"
objExcel_tny_master16721167666803.Range("input!M25").Value = "10"
objWorkbook_tny_master16721167666803.Save
objExcel_tny_master16721167666803.Application.DisplayAlerts = False
objExcel_tny_master16721167666803.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\tny_master16721167666803.xlsm'!GetData"
objExcel_tny_master16721167666803.Application.DisplayAlerts = False
objWorkbook_tny_master16721167666803.Save
objWorkbook_tny_master16721167666803.Close False 
objExcel_tny_master16721167666803.Application.Quit
Set objExcel_tny_master16721167666803 = Nothing 
Set objWorkbook_tny_master16721167666803 = Nothing
