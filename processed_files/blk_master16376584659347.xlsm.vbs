Set objExcel_blk_master16376584659347    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16376584659347 = objExcel_blk_master16376584659347.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16376584659347.xlsm")
objExcel_blk_master16376584659347.Range("input!C1").Value = "9"
objExcel_blk_master16376584659347.Range("input!J4").Value = "9000"
objExcel_blk_master16376584659347.Range("input!J3").Value = "0.9"
objExcel_blk_master16376584659347.Range("input!J5").Value = "17584.56"
objExcel_blk_master16376584659347.Range("input!J6").Value = "33"
objExcel_blk_master16376584659347.Range("input!J7").Value = "29"
objExcel_blk_master16376584659347.Range("input!J8").Value = "70"
objExcel_blk_master16376584659347.Range("input!J9").Value = "1.0124"
objExcel_blk_master16376584659347.Range("input!J10").Value = "25"
objExcel_blk_master16376584659347.Range("input!J11").Value = "28"
objExcel_blk_master16376584659347.Range("input!J12").Value = "10"
objWorkbook_blk_master16376584659347.Save
objExcel_blk_master16376584659347.Application.DisplayAlerts = False
objExcel_blk_master16376584659347.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16376584659347.xlsm'!GetData"
objExcel_blk_master16376584659347.Application.DisplayAlerts = False
objWorkbook_blk_master16376584659347.Save
objWorkbook_blk_master16376584659347.Close False 
objExcel_blk_master16376584659347.Application.Quit
Set objExcel_blk_master16376584659347 = Nothing 
Set objWorkbook_blk_master16376584659347 = Nothing
