Attribute VB_Name = "ScenarioTemplateRun"
Option Explicit

Sub runScenariosFromTable(mainTableName As String)
    Dim mainTable As zVLibTable
    Set mainTable = VLib.GetExcelTable(mainTableName, ThisWorkbook)
        
    Dim ScenarioTableNames() As String
    Dim LiveTableNames() As String
    
    Dim WorkbookPath As String
    Dim wb As Workbook
    
    Dim isOpen As Boolean
    
    Dim lr As ListRow, Active As Boolean, i As Long
    For Each lr In mainTable.ListObject.ListRows
    
        If ExistsInCollection(mainTable.ListObject.ListColumns, "Active") Then
            Active = mainTable.ColumnRange("Active")(lr.index).value
        Else
            Active = True ' default is true
        End If
        
        If Active Then
            ' Workbook schenanigans for if WorkbookPath column does not exist or value is empty
            If ExistsInCollection(mainTable.ListObject.ListColumns, "WorkbookPath") Then
                WorkbookPath = absolutePath(mainTable.ColumnRange("WorkbookPath")(lr.index).value)
            Else
                WorkbookPath = ""
            End If
            
            If "" & WorkbookPath = "" Then WorkbookPath = ThisWorkbook.FullName
            
            isOpen = VLib.IsWorkbookOpen(VLib.GetFilename(WorkbookPath))
            If isOpen Then
                Set wb = Workbooks(VLib.GetFilename(WorkbookPath))
            Else
                Set wb = Workbooks.Open(WorkbookPath)
            End If
            
            ' These tables must exist
            ScenarioTableNames = VLib.SplitTrim(mainTable.ColumnRange("ScenarioTableNames")(lr.index).value, vbLf)
            LiveTableNames = VLib.SplitTrim(mainTable.ColumnRange("LiveTableNames")(lr.index).value, vbLf)
            
            ' test that the lists are of equal length:
            If UBound(ScenarioTableNames) <> UBound(LiveTableNames) Then
                Application.GoTo mainTable.ColumnRange("ScenarioTableNames")(lr.index) ' select the cell with the error
                Err.Raise 514, , "Number of ScenarioTableNames and LiveTableNames should be the same"
            End If
            
            For i = LBound(ScenarioTableNames) To UBound(ScenarioTableNames)
                ' Do the scenario exporting
                exportScenarioTemplates ScenarioTableNames(i), LiveTableNames(i), wb
            Next i
            
            If Not isOpen Then
                wb.Close SaveChanges:=False
            End If
        End If
    Next lr
    
End Sub
