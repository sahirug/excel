Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\sahir\Documents\Macros\MacroWS.xlsm")

objExcel.Application.Run "MacroWS.xlsm!NewMacro2"
objWorkbook.Save
objExcel.DisplayAlerts = False 
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit
WScript.Quit