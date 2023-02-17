Set objExcel_blk_master16381728359773    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16381728359773 = objExcel_blk_master16381728359773.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16381728359773.xlsm")
objExcel_blk_master16381728359773.Range("input!C1").Value = "13"
objExcel_blk_master16381728359773.Range("input!J4").Value = "13000"
objExcel_blk_master16381728359773.Range("input!J3").Value = "0.9"
objExcel_blk_master16381728359773.Range("input!J5").Value = "15909.84"
objExcel_blk_master16381728359773.Range("input!J6").Value = "33"
objExcel_blk_master16381728359773.Range("input!J7").Value = "29"
objExcel_blk_master16381728359773.Range("input!J8").Value = "70"
objExcel_blk_master16381728359773.Range("input!J9").Value = "1.0124"
objExcel_blk_master16381728359773.Range("input!J10").Value = "20"
objExcel_blk_master16381728359773.Range("input!J11").Value = "28"
objExcel_blk_master16381728359773.Range("input!J12").Value = "10"
objWorkbook_blk_master16381728359773.Save
objExcel_blk_master16381728359773.Application.DisplayAlerts = False
objExcel_blk_master16381728359773.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16381728359773.xlsm'!GetData"
objExcel_blk_master16381728359773.Application.DisplayAlerts = False
objWorkbook_blk_master16381728359773.Save
objWorkbook_blk_master16381728359773.Close False 
objExcel_blk_master16381728359773.Application.Quit
Set objExcel_blk_master16381728359773 = Nothing 
Set objWorkbook_blk_master16381728359773 = Nothing
