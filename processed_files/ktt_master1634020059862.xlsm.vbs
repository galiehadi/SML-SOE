Set objExcel_ktt_master1634020059862    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master1634020059862 = objExcel_ktt_master1634020059862.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master1634020059862.xlsm")
objExcel_ktt_master1634020059862.Range("input!C1").Value = "75"
objExcel_ktt_master1634020059862.Range("input!J4").Value = "75000"
objExcel_ktt_master1634020059862.Range("input!J3").Value = "0.85"
objExcel_ktt_master1634020059862.Range("input!J5").Value = "16642.53"
objExcel_ktt_master1634020059862.Range("input!J6").Value = "34"
objExcel_ktt_master1634020059862.Range("input!J7").Value = "27"
objExcel_ktt_master1634020059862.Range("input!J8").Value = "63"
objExcel_ktt_master1634020059862.Range("input!J9").Value = "1.0123"
objExcel_ktt_master1634020059862.Range("input!J10").Value = "22"
objExcel_ktt_master1634020059862.Range("input!J11").Value = "27"
objExcel_ktt_master1634020059862.Range("input!J12").Value = "8"
objWorkbook_ktt_master1634020059862.Save
objExcel_ktt_master1634020059862.Application.DisplayAlerts = False
objExcel_ktt_master1634020059862.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master1634020059862.xlsm'!GetData"
objExcel_ktt_master1634020059862.Application.DisplayAlerts = False
objWorkbook_ktt_master1634020059862.Save
objWorkbook_ktt_master1634020059862.Close False 
objExcel_ktt_master1634020059862.Application.Quit
Set objExcel_ktt_master1634020059862 = Nothing 
Set objWorkbook_ktt_master1634020059862 = Nothing
