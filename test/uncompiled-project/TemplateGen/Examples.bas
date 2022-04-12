Attribute VB_Name = "Examples"
Option Explicit


' *********** In scenarios and live tables ***********
' REQUIRED columns in the scenarios and live tables:
'  - OutputTableNames
'    names of the Excel tables (ListObject) that needs to be exported
'    multiple names can be specified by listing them on new lines
'    (Alt + Enter in Excel cell)
'  - OutputFileNames
'    names of the output files to be exported
'    multiple names can be specified by listing them on new lines
'    (Alt + Enter in Excel cell)
'    NB: number of TableNames and FileNames must match
' OPTIONAL columns:
'  - RunMacros
'    New line separated (comma also supported) list of macros to run
'    (macros in the workbook where the Live and Scenarios tables live)
'  - Active
'    Whether to run the current line in a scenario table

' *********** In runScenarios table ***********
' REQUIRED columns
'  - ScenarioTableNames
'    Names of the scenario tables
'    multiple names can be specified by listing them on new lines
'    (Alt + Enter in Excel cell)
'  - LiveTableNames
'    Names of the live tables
'    multiple names can be specified by listing them on new lines
'    (Alt + Enter in Excel cell)
' OPTIONAL columns:
'  - WorkbookPath
'    Path to the workbook in which the Scenario and Live table lives
'    If ommitted or blank, defaults to ThisWorkbook
'  - Active
'    Whether to run the current line in a scenarios table

' example runs for the TemplateGenerator

Sub ExampleExportScenarioTemplates()
    
    ' runs the scenarios in the "scenarios" table using
    ' the "live" table as live values
    exportScenarioTemplates "scenarios", "live"
    
End Sub


Sub ExampleRunScenariosFromTable()
    
    ' runs all scenarios and live tables listed in the
    ' "ScenarioTableName" and "LiveTableName" tables
    runScenariosFromTable "runScenarios"
    
End Sub

