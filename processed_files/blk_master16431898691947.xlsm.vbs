Set objExcel_blk_master16431898691947    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16431898691947 = objExcel_blk_master16431898691947.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16431898691947.xlsm")
objExcel_blk_master16431898691947.Range("input!C1").Value = "9"
objExcel_blk_master16431898691947.Range("input!J4").Value = "9000"
objExcel_blk_master16431898691947.Range("input!J3").Value = "0.85653480857259"
objExcel_blk_master16431898691947.Range("input!J5").Value = "17584.56"
objExcel_blk_master16431898691947.Range("input!J6").Value = "25"
objExcel_blk_master16431898691947.Range("input!J7").Value = "25"
objExcel_blk_master16431898691947.Range("input!J8").Value = "55"
objExcel_blk_master16431898691947.Range("input!J9").Value = "1.0122"
objExcel_blk_master16431898691947.Range("input!J10").Value = "15"
objExcel_blk_master16431898691947.Range("input!J11").Value = "27"
objExcel_blk_master16431898691947.Range("input!J12").Value = "8"
objWorkbook_blk_master16431898691947.Save
objExcel_blk_master16431898691947.Application.DisplayAlerts = False
objExcel_blk_master16431898691947.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16431898691947.xlsm'!GetData"
objExcel_blk_master16431898691947.Application.DisplayAlerts = False
objWorkbook_blk_master16431898691947.Save
objWorkbook_blk_master16431898691947.Close False 
objExcel_blk_master16431898691947.Application.Quit
Set objExcel_blk_master16431898691947 = Nothing 
Set objWorkbook_blk_master16431898691947 = Nothing
