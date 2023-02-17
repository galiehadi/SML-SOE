Set objExcel_blk_master16376580183035    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16376580183035 = objExcel_blk_master16376580183035.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16376580183035.xlsm")
objExcel_blk_master16376580183035.Range("input!C1").Value = "9"
objExcel_blk_master16376580183035.Range("input!J4").Value = "9000"
objExcel_blk_master16376580183035.Range("input!J3").Value = "0.85653480857259"
objExcel_blk_master16376580183035.Range("input!J5").Value = "17584.56"
objExcel_blk_master16376580183035.Range("input!J6").Value = "25"
objExcel_blk_master16376580183035.Range("input!J7").Value = "25"
objExcel_blk_master16376580183035.Range("input!J8").Value = "55"
objExcel_blk_master16376580183035.Range("input!J9").Value = "1.0122"
objExcel_blk_master16376580183035.Range("input!J10").Value = "15"
objExcel_blk_master16376580183035.Range("input!J11").Value = "27"
objExcel_blk_master16376580183035.Range("input!J12").Value = "8"
objWorkbook_blk_master16376580183035.Save
objExcel_blk_master16376580183035.Application.DisplayAlerts = False
objExcel_blk_master16376580183035.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16376580183035.xlsm'!GetData"
objExcel_blk_master16376580183035.Application.DisplayAlerts = False
objWorkbook_blk_master16376580183035.Save
objWorkbook_blk_master16376580183035.Close False 
objExcel_blk_master16376580183035.Application.Quit
Set objExcel_blk_master16376580183035 = Nothing 
Set objWorkbook_blk_master16376580183035 = Nothing
