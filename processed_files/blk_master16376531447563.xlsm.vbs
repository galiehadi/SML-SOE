Set objExcel_blk_master16376531447563    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16376531447563 = objExcel_blk_master16376531447563.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16376531447563.xlsm")
objExcel_blk_master16376531447563.Range("input!C1").Value = "13"
objExcel_blk_master16376531447563.Range("input!J4").Value = "13000"
objExcel_blk_master16376531447563.Range("input!J3").Value = "0.9"
objExcel_blk_master16376531447563.Range("input!J5").Value = "17584.56"
objExcel_blk_master16376531447563.Range("input!J6").Value = "33"
objExcel_blk_master16376531447563.Range("input!J7").Value = "29"
objExcel_blk_master16376531447563.Range("input!J8").Value = "70"
objExcel_blk_master16376531447563.Range("input!J9").Value = "1.0124"
objExcel_blk_master16376531447563.Range("input!J10").Value = "25"
objExcel_blk_master16376531447563.Range("input!J11").Value = "28"
objExcel_blk_master16376531447563.Range("input!J12").Value = "10"
objWorkbook_blk_master16376531447563.Save
objExcel_blk_master16376531447563.Application.DisplayAlerts = False
objExcel_blk_master16376531447563.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16376531447563.xlsm'!GetData"
objExcel_blk_master16376531447563.Application.DisplayAlerts = False
objWorkbook_blk_master16376531447563.Save
objWorkbook_blk_master16376531447563.Close False 
objExcel_blk_master16376531447563.Application.Quit
Set objExcel_blk_master16376531447563 = Nothing 
Set objWorkbook_blk_master16376531447563 = Nothing
