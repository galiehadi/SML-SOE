Set objExcel_amg_master1638331192485    = CreateObject("Excel.Application")
Set objWorkbook_amg_master1638331192485 = objExcel_amg_master1638331192485.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master1638331192485.xlsm")
objExcel_amg_master1638331192485.Range("input!C1").Value = "22"
objExcel_amg_master1638331192485.Range("input!J4").Value = "22000"
objExcel_amg_master1638331192485.Range("input!J3").Value = "0.9"
objExcel_amg_master1638331192485.Range("input!J5").Value = "16328.52"
objExcel_amg_master1638331192485.Range("input!J6").Value = "33"
objExcel_amg_master1638331192485.Range("input!J7").Value = "29"
objExcel_amg_master1638331192485.Range("input!J8").Value = "65"
objExcel_amg_master1638331192485.Range("input!J9").Value = "1.0124"
objExcel_amg_master1638331192485.Range("input!J10").Value = "20"
objExcel_amg_master1638331192485.Range("input!J11").Value = "29"
objExcel_amg_master1638331192485.Range("input!J12").Value = "10"
objWorkbook_amg_master1638331192485.Save
objExcel_amg_master1638331192485.Application.DisplayAlerts = False
objExcel_amg_master1638331192485.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master1638331192485.xlsm'!GetData"
objExcel_amg_master1638331192485.Application.DisplayAlerts = False
objWorkbook_amg_master1638331192485.Save
objWorkbook_amg_master1638331192485.Close False 
objExcel_amg_master1638331192485.Application.Quit
Set objExcel_amg_master1638331192485 = Nothing 
Set objWorkbook_amg_master1638331192485 = Nothing
