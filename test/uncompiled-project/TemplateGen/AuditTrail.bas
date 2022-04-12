Attribute VB_Name = "AuditTrail"
Option Explicit

Function generateAuditItem(live As zVLibTable, scenarios As zVLibTable, abspath As String, templateName As String) As Dictionary

    Dim fso As New FileSystemObject
    Dim wb As Workbook
    Set wb = scenarios.ListObject.Range.Worksheet.Parent
    
    Dim item As Dictionary
    Dim parameters As Dictionary
    Dim additional As Dictionary
    Dim i As Long
    
    Set item = New Dictionary
    Set parameters = New Dictionary
    Set additional = New Dictionary
    
    Dim modified As Variant
    modified = fso.GetFile(abspath).DateLastModified
    
    
    item.Add "name", VLib.GetFilename(abspath)
    item.Add "size", FileLen(abspath)
    item.Add "md5_hash", ComputeMd5(abspath)
    'item.Add "modified", (modified * 86400) - 2209168804# 'TODO: get correct calculation for serial+timezone to unix epoch time
    item.Add "parameters", parameters
    item.Add "additional", additional
    
    parameters.Add "scenarios_table", scenarios.Name
    parameters.Add "live_table", live.Name
    
    With live.ListObject
        For i = 1 To .Range.Columns.count
            parameters.Add .Range(1, i).value, .Range(2, i).value
        Next i
    End With
    
    additional.Add "modified", Format(modified, "yyyy-mm-dd hh:mm:ss.0000")
    additional.Add "path", abspath
    additional.Add "workbook", live.ListObject.Range.Worksheet.Parent.FullName
    ' the table within the workbook
    additional.Add "table", templateName
    'additional.Add "md5_hash", ComputeMd5(abspath)
    
    Set generateAuditItem = item
End Function


Sub extendAuditJson(auditDicts As Dictionary, item As Dictionary)
    Dim fso As New FileSystemObject
    Dim dirpath As String
    Dim jsonpath As String
    Dim auditDict As Dictionary
    Dim strFileContent As String
    Dim iFile As Integer
    Dim oFile As Integer
    Dim jsonstr As String
    
    
    dirpath = LCase( _
                fso.GetParentFolderName( _
                  fso.GetAbsolutePathName(item.item("additional").item("path"))))
                  
    jsonpath = dirpath & "\" & ".testaudit.json"
    
    If auditDicts.Exists(jsonpath) Then
        Set auditDict = auditDicts.item(jsonpath)
        
    ElseIf VLib.FileExists(jsonpath) Then
        iFile = FreeFile
        Open jsonpath For Input As #iFile
            jsonstr = Input(LOF(iFile), iFile)
        Close #iFile
        If Trim(jsonstr) = "" Then jsonstr = "{}"
        Set auditDict = Json.ParseJson(jsonstr)
        
        auditDicts.Add jsonpath, auditDict
        
    Else
        Set auditDict = New Dictionary
        auditDicts.Add jsonpath, auditDict
       
    End If
    
    If auditDict.Exists(item.item("name")) Then
        Set auditDict.item(item.item("name")) = item
    Else
        auditDict.Add item.item("name"), item
    End If
    
    
End Sub
    

Sub ExportDicts(auditDicts As Dictionary)
    Dim auditDict As Dictionary
    Dim k As Variant
    Dim jsonstr As String
    Dim jsonpath As String
    Dim oFile As Integer
    For Each k In auditDicts.keys()
        Set auditDict = auditDicts(k)
        jsonstr = Json.ConvertToJson(auditDict, 4)
        oFile = FreeFile
        jsonpath = k
        Open jsonpath For Output As #oFile
            Print #oFile, jsonstr
        Close #oFile
        
    Next k
End Sub

