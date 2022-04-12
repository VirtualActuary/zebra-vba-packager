Attribute VB_Name = "ScenarioTemplateGeneration"
Option Explicit

Sub exportScenarioTemplates(ScenarioTableName As String, _
                            LiveTableName As String, _
                            Optional wb As Workbook = Nothing)
                            
                            
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim fso As New FileSystemObject
    
    Dim auditDicts As Dictionary
    Set auditDicts = New Dictionary
    
    Dim scenarios As zVLibTable
    Dim live As zVLibTable
    Dim template As zVLibTable
    Dim abspath As String
    Dim RunMacros As String
    Dim macro As Variant
        
    
    Dim csvout As zWsCsvInterface
    Set csvout = New zWsCsvInterface
    With csvout.parseConfig
        .fieldsDelimiter = ","
        .recordsDelimiter = vbCrLf
        .headers = True
    End With
    
    
    Set scenarios = VLib.GetExcelTable(ScenarioTableName, wb)
    Set live = VLib.GetExcelTable(LiveTableName, wb)
    
    Dim lr As ListRow, Active As Boolean
    Dim OutputFileNames() As String, OutputTableNames() As String, i As Long
    
    For Each lr In scenarios.ListObject.ListRows
        
        If ExistsInCollection(scenarios.ListObject.ListColumns, "Active") Then
            Active = scenarios.ColumnRange("Active")(lr.index).value
        Else
            Active = True ' default is true
        End If
        
        If Active Then
        
            live.ListObject.ListRows(1).Range.value = lr.Range.value
        
            ' Collect macros to run, and run them
            If ExistsInCollection(live.ListObject.ListColumns, "RunMacros") Then
                RunMacros = live.ColumnRange("RunMacros").value
            Else
                RunMacros = ""
            End If
            
            ' support both comma and new line delimited macros
            RunMacros = Replace(RunMacros, ",", vbLf)
            
            For Each macro In VLib.SplitTrim(RunMacros, vbLf)
                Application.Run "'" & wb.Name & "'!" & macro
            Next macro
            
            OutputTableNames = VLib.SplitTrim(live.ColumnRange("OutputTableNames").value, vbLf)
            OutputFileNames = VLib.SplitTrim(live.ColumnRange("OutputFileNames").value, vbLf)
            
            If UBound(OutputFileNames) <> UBound(OutputTableNames) Then
                Application.GoTo scenarios.ColumnRange("OutputTableNames")(lr.index) ' select the cell with the error
                Err.Raise 514, , "Number of OutputTableNames and OutputFileNames should be the same"
            End If
            
            For i = LBound(OutputFileNames) To UBound(OutputFileNames)
            
                abspath = absolutePath(OutputFileNames(i), wb)
                
                Set template = VLib.GetExcelTable(OutputTableNames(i), wb)
                Application.Calculate
                With csvout
                    ' ExportToCSV appends to an existing CSV file, so we need to delete the file first
                    If VLib.FileExists(abspath) Then fso.DeleteFile abspath
                    ' Ensure the directory exists
                    VLib.MkDirRecursive VLib.GetDirectoryName(abspath)
                    .parseConfig.Path = abspath
                    .ExportToCSV applyCsvTransformations(template.ListObject.Range.value)
                    If .exportSuccess = False Then
                        Err.Raise 6, , "CSV write to " & abspath & " failed, check file permissions or if open in another application."
                    End If
                End With
                
                extendAuditJson auditDicts, generateAuditItem(live, scenarios, abspath, OutputTableNames(i))
            Next i
            
        End If

    Next lr
    
    ' export all at once for runtime efficiency
    ExportDicts auditDicts
            
End Sub



