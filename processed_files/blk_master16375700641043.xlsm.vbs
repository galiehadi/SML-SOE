Set objExcel_blk_master16375700641043    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16375700641043 = objExcel_blk_master16375700641043.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16375700641043.xlsm")
objExcel_blk_master16375700641043.Range("input!C1").Value = "0"
objExcel_blk_master16375700641043.Range("input!J4").Value = "0"
objExcel_blk_master16375700641043.Range("input!J3").Value = "0"
objExcel_blk_master16375700641043.Range("input!J5").Value = "0"
objExcel_blk_master16375700641043.Range("input!J6").Value = "0"
objExcel_blk_master16375700641043.Range("input!J7").Value = "0"
objExcel_blk_master16375700641043.Range("input!J8").Value = "0"
objExcel_blk_master16375700641043.Range("input!J9").Value = "0"
objExcel_blk_master16375700641043.Range("input!J10").Value = "0"
objExcel_blk_master16375700641043.Range("input!J11").Value = "0"
objExcel_blk_master16375700641043.Range("input!J12").Value = "0"
objWorkbook_blk_master16375700641043.Save
objExcel_blk_master16375700641043.Application.DisplayAlerts = False
objExcel_blk_master16375700641043.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16375700641043.xlsm'!GetData"
objExcel_blk_master16375700641043.Application.DisplayAlerts = False
objWorkbook_blk_master16375700641043.Save
objWorkbook_blk_master16375700641043.Close False 
objExcel_blk_master16375700641043.Application.Quit
Set objExcel_blk_master16375700641043 = Nothing 
Set objWorkbook_blk_master16375700641043 = Nothing
