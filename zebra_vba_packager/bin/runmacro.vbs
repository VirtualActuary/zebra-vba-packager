Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("WScript.Shell")

ws = fso.GetAbsolutePathName(WScript.Arguments(0))
macro = ""
if WScript.Arguments.Count > 1 then
    macro = WScript.Arguments(1)
end if

Set XLS = CreateObject("Excel.application")

objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\VBAWarnings", 1, "REG_DWORD"

Set xlBook = XLS.Workbooks.Open(ws, 0, False)

if macro <> "" then
    xlBook.Application.run macro
end if

XLS.DisplayAlerts = False
xlBook.Save
XLS.Quit
