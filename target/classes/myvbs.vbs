path = WScript.Arguments.Item(0)
fileName = WScript.Arguments.Item(1)
macroName = WScript.Arguments.Item(2)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(path)

' objExcel.Application.Run "MacroWS.xlsm!NewMacro2"
objExcel.Application.Run fileName + "!" + macroName
objWorkbook.Save
objExcel.DisplayAlerts = False 
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit
WScript.Quit