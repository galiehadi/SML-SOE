Set objExcel_blk_master16375653368034    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16375653368034 = objExcel_blk_master16375653368034.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16375653368034.xlsm")
objExcel_blk_master16375653368034.Range("input!C1").Value = "15"
objExcel_blk_master16375653368034.Range("input!J4").Value = "15000"
objExcel_blk_master16375653368034.Range("input!J3").Value = "0.85653480857259"
objExcel_blk_master16375653368034.Range("input!J5").Value = "17584.56"
objExcel_blk_master16375653368034.Range("input!J6").Value = "25"
objExcel_blk_master16375653368034.Range("input!J7").Value = "25"
objExcel_blk_master16375653368034.Range("input!J8").Value = "55"
objExcel_blk_master16375653368034.Range("input!J9").Value = "1.0122"
objExcel_blk_master16375653368034.Range("input!J10").Value = "15"
objExcel_blk_master16375653368034.Range("input!J11").Value = "27"
objExcel_blk_master16375653368034.Range("input!J12").Value = "8"
objWorkbook_blk_master16375653368034.Save
objExcel_blk_master16375653368034.Application.DisplayAlerts = False
objExcel_blk_master16375653368034.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16375653368034.xlsm'!GetData"
objExcel_blk_master16375653368034.Application.DisplayAlerts = False
objWorkbook_blk_master16375653368034.Save
objWorkbook_blk_master16375653368034.Close False 
objExcel_blk_master16375653368034.Application.Quit
Set objExcel_blk_master16375653368034 = Nothing 
Set objWorkbook_blk_master16375653368034 = Nothing
