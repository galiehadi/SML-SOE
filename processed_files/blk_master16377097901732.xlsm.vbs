Set objExcel_blk_master16377097901732    = CreateObject("Excel.Application")
Set objWorkbook_blk_master16377097901732 = objExcel_blk_master16377097901732.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\blk_master16377097901732.xlsm")
objExcel_blk_master16377097901732.Range("input!C1").Value = ""
objExcel_blk_master16377097901732.Range("input!J4").Value = "0"
objExcel_blk_master16377097901732.Range("input!J3").Value = ""
objExcel_blk_master16377097901732.Range("input!J5").Value = "0"
objExcel_blk_master16377097901732.Range("input!J6").Value = ""
objExcel_blk_master16377097901732.Range("input!J7").Value = ""
objExcel_blk_master16377097901732.Range("input!J8").Value = ""
objExcel_blk_master16377097901732.Range("input!J9").Value = ""
objExcel_blk_master16377097901732.Range("input!J10").Value = ""
objExcel_blk_master16377097901732.Range("input!J11").Value = ""
objExcel_blk_master16377097901732.Range("input!J12").Value = ""
objWorkbook_blk_master16377097901732.Save
objExcel_blk_master16377097901732.Application.DisplayAlerts = False
objExcel_blk_master16377097901732.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\blk_master16377097901732.xlsm'!GetData"
objExcel_blk_master16377097901732.Application.DisplayAlerts = False
objWorkbook_blk_master16377097901732.Save
objWorkbook_blk_master16377097901732.Close False 
objExcel_blk_master16377097901732.Application.Quit
Set objExcel_blk_master16377097901732 = Nothing 
Set objWorkbook_blk_master16377097901732 = Nothing
