Set objExcel_amg_master16389432711647    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16389432711647 = objExcel_amg_master16389432711647.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16389432711647.xlsm")
objExcel_amg_master16389432711647.Range("input!C1").Value = "24.5"
objExcel_amg_master16389432711647.Range("input!J4").Value = "24500"
objExcel_amg_master16389432711647.Range("input!J3").Value = "0.9"
objExcel_amg_master16389432711647.Range("input!J5").Value = "17165.88"
objExcel_amg_master16389432711647.Range("input!J6").Value = "33"
objExcel_amg_master16389432711647.Range("input!J7").Value = "29"
objExcel_amg_master16389432711647.Range("input!J8").Value = "65"
objExcel_amg_master16389432711647.Range("input!J9").Value = "1.0124"
objExcel_amg_master16389432711647.Range("input!J10").Value = "20"
objExcel_amg_master16389432711647.Range("input!J11").Value = "29"
objExcel_amg_master16389432711647.Range("input!J12").Value = "10"
objWorkbook_amg_master16389432711647.Save
objExcel_amg_master16389432711647.Application.DisplayAlerts = False
objExcel_amg_master16389432711647.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16389432711647.xlsm'!GetData"
objExcel_amg_master16389432711647.Application.DisplayAlerts = False
objWorkbook_amg_master16389432711647.Save
objWorkbook_amg_master16389432711647.Close False 
objExcel_amg_master16389432711647.Application.Quit
Set objExcel_amg_master16389432711647 = Nothing 
Set objWorkbook_amg_master16389432711647 = Nothing