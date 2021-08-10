Set fso = Createobject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("WScript.Shell")

'Get args
Excelfile = fso.GetAbsolutePathName(WScript.Arguments(0))
ExcelDest = fso.GetAbsolutePathName(WScript.Arguments(1))

'# Disable macro security (yes, this is very naughty indeed)
set XLS = CreateObject("Excel.Application")
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\VBAWarnings", 1, "REG_DWORD"

XLS.EnableEvents = False
XLS.DisplayAlerts = False
XLS.Visible = False
XLS.ScreenUpdating = False
XLS.UserControl = False
XLS.Interactive = False

' Open Excel Workbook 
Set WB = XLS.Workbooks.Open(ExcelFile)

WB.SaveAs ExcelDest, 51 '51 = xlOpenXMLWorkbook xlsx
WB.Close
XLS.Quit
