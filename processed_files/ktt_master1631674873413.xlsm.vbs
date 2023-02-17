Set objExcel_ktt_master1631674873413    = CreateObject("Excel.Application")
Set objWorkbook_ktt_master1631674873413 = objExcel_ktt_master1631674873413.Workbooks.Open("C:\xampp\htdocs\online_elink\processed_files\ktt_master1631674873413.xlsm")
objExcel_ktt_master1631674873413.Range("input!C1").Value = "75"
objExcel_ktt_master1631674873413.Range("input!J4").Value = "75000"
objExcel_ktt_master1631674873413.Range("input!J3").Value = "0.85653480857259"
objExcel_ktt_master1631674873413.Range("input!J5").Value = "17584.56"
objExcel_ktt_master1631674873413.Range("input!J6").Value = "25"
objExcel_ktt_master1631674873413.Range("input!J7").Value = "25"
objExcel_ktt_master1631674873413.Range("input!J8").Value = "55"
objExcel_ktt_master1631674873413.Range("input!J9").Value = "1.0122"
objExcel_ktt_master1631674873413.Range("input!J10").Value = "15"
objExcel_ktt_master1631674873413.Range("input!J11").Value = "27"
objExcel_ktt_master1631674873413.Range("input!J12").Value = "8"
objWorkbook_ktt_master1631674873413.Save
objExcel_ktt_master1631674873413.Application.DisplayAlerts = False
objExcel_ktt_master1631674873413.Application.Run "'C:\xampp\htdocs\online_elink\processed_files\ktt_master1631674873413.xlsm'!GetData"
objExcel_ktt_master1631674873413.Application.DisplayAlerts = False
objWorkbook_ktt_master1631674873413.Save
objWorkbook_ktt_master1631674873413.Close False 
objExcel_ktt_master1631674873413.Application.Quit
Set objExcel_ktt_master1631674873413 = Nothing 
Set objWorkbook_ktt_master1631674873413 = Nothing
