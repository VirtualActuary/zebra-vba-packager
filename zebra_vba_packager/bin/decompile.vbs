Set fso = Createobject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("WScript.Shell")
Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)

Sub Print(input)
    On Error Resume Next
    If not StdOut is Nothing Then
        StdOut.Write input
    End If
End Sub

Function RandString(strLen)
    Dim str
    Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
    strLen=len(LETTERS)
    Randomize
    For i = 1 to strLen
        str = str & Mid(LETTERS, Int((Len(LETTERS)-1+1)*Rnd+1), 1)
    Next
    RandString = str
End Function

Function EnsureDir(DirName)
    If not fso.FolderExists(DirName) Then
        fso.CreateFolder DirName
    End If
End Function

'Get all the filenames in order'
Excelfile = fso.GetAbsolutePathName(WScript.Arguments(0))
if not fso.FileExists(Excelfile) then
  Err.Raise vbObjectError+10, "", "File not found: " & Excelfile
end if

FDir = fso.GetParentFolderName(Excelfile)
FName = fso.GetBaseName(Excelfile)
FExt = fso.getextensionname(Excelfile)

RandName = RandString(20)
ExcelCopy = fso.GetSpecialFolder(2) & "\" & Fname & RandName & "." & FExt 'hide somewhere in temp dir
MacroDir = FDir & "\" & FName
If WScript.Arguments.Count > 1 Then
	MacroDir = fso.GetAbsolutePathName(WScript.Arguments(1))
End If

MacroDir1   = MacroDir & "\Modules"
MacroDir2   = MacroDir & "\ClassModules"
MacroDir3   = MacroDir & "\Forms"
MacroDir100 = MacroDir & "\ExcelObjects"

ExcelPreDest = fso.GetSpecialFolder(2) & "\" & Fname & RandName & ".xlsx"
ExcelDest = MacroDir & "\" & FName & ".xlsx"

'Call script to remove password from the excel file
scriptdir=fso.GetParentFolderName(WScript.ScriptFullName)

fso.CopyFile ExcelFile, ExcelCopy
Do until fso.FileExists(ExcelCopy)
    WScript.Sleep(20)
Loop

' Decompiling
'# Disable macro security (yes, this is very naughty indeed)
set XLS = CreateObject("Excel.Application")

' Disable Macro guards in the registry
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\VBAWarnings", 1, "REG_DWORD"

XLS.EnableEvents = False
XLS.DisplayAlerts = False
XLS.Visible = False
XLS.ScreenUpdating = False
XLS.UserControl = False
XLS.Interactive = False

' Open Excel Workbook 
Set WB = XLS.Workbooks.Open(ExcelCopy)

'Test to see if the objects are accecible
On Error Resume Next
    With WB.VBProject
        For i = 1 To .VBComponents.Count
            .VBComponents(i)
        Next
    End With

'If decompile-able
If Err.Number <> 0 Then
    EnsureDir(MacroDir)
    WScript.Sleep(20)
    fso.DeleteFolder(MacroDir)
    WScript.Sleep(20)
    EnsureDir(MacroDir)

    'Reset error handling'
    On Error Goto 0

    With WB.VBProject

        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Then

                EnsureDir(MacroDir1)
                .VBComponents(i).Export MacroDir1 & "\" & .VBComponents(i).CodeModule.Name & ".bas"

            ElseIf .VBComponents(i).Type = 2 Then

                EnsureDir(MacroDir2)
                .VBComponents(i).Export MacroDir2 & "\" & .VBComponents(i).CodeModule.Name & ".cls"

            ElseIf .VBComponents(i).Type = 3 Then

                EnsureDir(MacroDir3)
                .VBComponents(i).Export MacroDir3 & "\" & .VBComponents(i).CodeModule.Name & ".frm"

            'We have a special case for the ThisWorkbook Object
            ElseIf .VBComponents(i).Type = 100 Then
                If .VBComponents(i).Name = "ThisWorkbook" Then
                        CodeName = "ThisWorkbook"
                    Else
                        CodeName = .VBComponents(i).Properties("Name").Value
                End If
                CodeName = lcase(CodeName)

                EnsureDir(MacroDir100)
                IF .VBComponents(i).CodeModule.CountOfLines Then
                    Set oFile = fso.CreateTextFile(MacroDir100 & "\" & CodeName & ".txt")
                    oFile.Write( _
                        .VBComponents(i).CodeModule.Lines(1, _
                           .VBComponents(i).CodeModule.CountOfLines _
                           ) _
                        )
                    oFile.Close
                End If

            End If
        Next
    End with

Else
    Err.Raise 507, "Could not Access VBComponents, did not decompile project!"
End If

'Reset error handling'
On Error Goto 0

WB.SaveAs ExcelPreDest, 51 '51 = xlOpenXMLWorkbook xlsx
WB.Close
XLS.Quit


fso.CopyFile ExcelPreDest, ExcelDest
Do until fso.FileExists(ExcelDest)
    WScript.Sleep(20)
Loop

fso.DeleteFile(ExcelCopy)
fso.DeleteFile(ExcelPreDest)
