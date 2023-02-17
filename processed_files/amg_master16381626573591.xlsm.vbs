Set objExcel_amg_master16381626573591    = CreateObject("Excel.Application")
Set objWorkbook_amg_master16381626573591 = objExcel_amg_master16381626573591.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\amg_master16381626573591.xlsm")
objExcel_amg_master16381626573591.Range("input!C1").Value = "15"
objExcel_amg_master16381626573591.Range("input!J4").Value = "15000"
objExcel_amg_master16381626573591.Range("input!J3").Value = "0.88"
objExcel_amg_master16381626573591.Range("input!J5").Value = "16747.2"
objExcel_amg_master16381626573591.Range("input!J6").Value = "33"
objExcel_amg_master16381626573591.Range("input!J7").Value = "28"
objExcel_amg_master16381626573591.Range("input!J8").Value = "67"
objExcel_amg_master16381626573591.Range("input!J9").Value = "1.0124"
objExcel_amg_master16381626573591.Range("input!J10").Value = "20"
objExcel_amg_master16381626573591.Range("input!J11").Value = "30"
objExcel_amg_master16381626573591.Range("input!J12").Value = "9"
objWorkbook_amg_master16381626573591.Save
objExcel_amg_master16381626573591.Application.DisplayAlerts = False
objExcel_amg_master16381626573591.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\amg_master16381626573591.xlsm'!GetData"
objExcel_amg_master16381626573591.Application.DisplayAlerts = False
objWorkbook_amg_master16381626573591.Save
objWorkbook_amg_master16381626573591.Close False 
objExcel_amg_master16381626573591.Application.Quit
Set objExcel_amg_master16381626573591 = Nothing 
Set objWorkbook_amg_master16381626573591 = Nothing
