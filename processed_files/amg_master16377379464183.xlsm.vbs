Set objExcel_amg_master16377379464183    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16377379464183 = objExcel_amg_master16377379464183.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16377379464183.xlsm")
objExcel_amg_master16377379464183.Range("input!C1").Value = ""
objExcel_amg_master16377379464183.Range("input!J4").Value = "0"
objExcel_amg_master16377379464183.Range("input!J3").Value = ""
objExcel_amg_master16377379464183.Range("input!J5").Value = "0"
objExcel_amg_master16377379464183.Range("input!J6").Value = ""
objExcel_amg_master16377379464183.Range("input!J7").Value = ""
objExcel_amg_master16377379464183.Range("input!J8").Value = ""
objExcel_amg_master16377379464183.Range("input!J9").Value = ""
objExcel_amg_master16377379464183.Range("input!J10").Value = ""
objExcel_amg_master16377379464183.Range("input!J11").Value = ""
objExcel_amg_master16377379464183.Range("input!J12").Value = ""
objWorkbook_amg_master16377379464183.Save
objExcel_amg_master16377379464183.Application.DisplayAlerts = False
objExcel_amg_master16377379464183.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16377379464183.xlsm'!GetData"
objExcel_amg_master16377379464183.Application.DisplayAlerts = False
objWorkbook_amg_master16377379464183.Save
objWorkbook_amg_master16377379464183.Close False 
objExcel_amg_master16377379464183.Application.Quit
Set objExcel_amg_master16377379464183 = Nothing 
Set objWorkbook_amg_master16377379464183 = Nothing
