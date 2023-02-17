Set objExcel_blk_master16497455665814    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16497455665814 = objExcel_blk_master16497455665814.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16497455665814.xlsm")
objExcel_blk_master16497455665814.Range("input!C1").Value = "8"
objExcel_blk_master16497455665814.Range("input!J4").Value = "8000"
objExcel_blk_master16497455665814.Range("input!J3").Value = "0.8"
objExcel_blk_master16497455665814.Range("input!J5").Value = "14653.8"
objExcel_blk_master16497455665814.Range("input!J6").Value = "25"
objExcel_blk_master16497455665814.Range("input!J7").Value = "25"
objExcel_blk_master16497455665814.Range("input!J8").Value = "55"
objExcel_blk_master16497455665814.Range("input!J9").Value = "1.0122"
objExcel_blk_master16497455665814.Range("input!J10").Value = "15"
objExcel_blk_master16497455665814.Range("input!J11").Value = "27"
objExcel_blk_master16497455665814.Range("input!J12").Value = "8"
objWorkbook_blk_master16497455665814.Save
objExcel_blk_master16497455665814.Application.DisplayAlerts = False
objExcel_blk_master16497455665814.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16497455665814.xlsm'!GetData"
objExcel_blk_master16497455665814.Application.DisplayAlerts = False
objWorkbook_blk_master16497455665814.Save
objWorkbook_blk_master16497455665814.Close False 
objExcel_blk_master16497455665814.Application.Quit
Set objExcel_blk_master16497455665814 = Nothing 
Set objWorkbook_blk_master16497455665814 = Nothing
