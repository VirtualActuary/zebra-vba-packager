'https://gist.github.com/coldfusion39/8f7e6bd6721514e01da5'

'Get all the filenames in order'
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

function getFiles(dirname)
    Set col = CreateObject("Scripting.Dictionary")
    getFiles_recursivehelper fso.GetFolder(dirname), col
    Set getFiles = col

end function

sub getFiles_recursivehelper(Folder, col)
  For Each File in Folder.Files
    col.Add col.Count, File
  Next

  For Each Subfolder in Folder.SubFolders
    getFiles_recursivehelper Subfolder, col
  Next
end sub

RandName = RandString(20)
MacroDir =  fso.GetAbsolutePathName(WScript.Arguments(0)) 'Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-1)

' Remove possible trailing "\"
IF Right(MacroDir, 1) = "\" then
    MacroDir = Left(MacroDir, LEN(MacroDir)-1)
End If

FDir = fso.GetParentFolderName(MacroDir)
FName = fso.GetBaseName(MacroDir)

Excel = MacroDir & "\" & FName & ".xlsx"
ExcelCopy = fso.GetSpecialFolder(2) &  "\" & FName & RandName & ".xlsx"
ExcelOutCopy = fso.GetSpecialFolder(2) &  "\" & FName & RandName & "--compiled.xlsb"
ExcelOut = FDir & "\" & FName & ".xlsb"

'Start with compiling the project
If fso.FileExists(ExcelOut) Then
    result = MsgBox ("File already exists:" & vbCrLf & ExcelOut & vbCrLf & vbCrLf & "Do you want to replace it?", _
                     vbYesNo + vbQuestion, "Confirm Compile")
    If result = vbNo Then wscript.quit
End If


'Print ">> Compiling: " & MacroDir & vbCrLf

'Print ">> Copy "& Excel &" to %TMP% (avoid Excel locks)" & vbCrLf
fso.CopyFile Excel, ExcelCopy
Do until fso.FileExists(ExcelCopy)
    WScript.Sleep(20)
Loop

'Print ">> Open Excel.exe background process" & vbCrLf

'# Create Excel objects
'Add-Type -AssemblyName Microsoft.Office.Interop.Excel
set XLS = CreateObject("Excel.Application")

'# Disable macro security
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & XLS.Version & "\Excel\Security\VBAWarnings", 1, "REG_DWORD"

XLS.EnableEvents = False 
XLS.DisplayAlerts = False
XLS.Visible = False
XLS.ScreenUpdating = False
XLS.UserControl = False
XLS.Interactive = False

'Print vbCrLf
'Print ">> Open Excel Workbook " & ExcelCopy & vbCrLf
Set Workbook = XLS.Workbooks.Open(ExcelCopy)

I = 0
For Each oFile in getFiles(MacroDir).Items
    ext       = Mid(oFile.Name,len(oFile.Name)-3,4)
    sheetname = Mid(oFile.Name,1,len(oFile.Name)-4)

    If ext = ".cls" or ext = ".bas" or ext = ".frm" or ext = ".txt" Then
        I = I + 1

        'We have a special case for txt Object, we need to find their
        If ext = ".txt" Then
            With Workbook.VBProject
                For J = 1 To .VBComponents.Count 
                    
                    If .VBComponents(j).Name = "ThisWorkbook" Then
                        CodeName = "ThisWorkbook"
                    Else
                        CodeName = .VBComponents(j).Properties("Name").Value
                    End If
                    CodeName = lcase(CodeName)

                    If StrComp(CodeName, sheetname, 1) = 0 Then
                        .VBComponents(J).CodeModule.DeleteLines _
                            1, .VBComponents(J).CodeModule.CountOfLines
                        .VBComponents(J).CodeModule.AddFromFile(oFile.Path)
                        Exit For
                    End If
                Next
            End With 
            if StrComp(CodeName, sheetname, 1) <> 0 then
                Err.Raise 507, "Failed to inject VBA module for sheet " & sheetname & "; no such sheet!"
            end If
        Else
            Workbook.VBProject.VBComponents.Import(oFile.Path)
        End If
    End If
Next 

'# Save the document
Workbook.SaveAs ExcelOutCopy, 50 '50 = xlExcel12 (Excel Binary Workbook in 2007-2016 with or without macro's, xlsb) '

'Print ">> Save As " & ExcelOutCopy & vbCrLf
'Print ">> Close Excel " & vbCrLf
Workbook.Close 
XLS.Quit

'Print ">> Copy to " & ExcelOut & vbCrLf
fso.CopyFile ExcelOutCopy, ExcelOut
Do until fso.FileExists(ExcelOut)
    WScript.Sleep(20)
Loop

fso.DeleteFile(ExcelCopy)
fso.DeleteFile(ExcelOutCopy)