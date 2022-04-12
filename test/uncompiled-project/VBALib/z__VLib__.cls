VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "z__VLib__"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory"  (dest As Any, source As Any, ByVal bytes As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory"  (dest As Any, source As Any, ByVal bytes As Long)
#End If
#If VBA7 Then
    Private Declare PtrSafe Function CallNamedPipe Lib "kernel32"  Alias "CallNamedPipeA" (  ByVal lpNamedPipeName As String,  ByVal lpInBuffer As Any, ByVal nInBufferSize As LongPtr,  ByRef lpOutBuffer As Any, ByVal nOutBufferSize As LongPtr,  ByRef lpBytesRead As LongPtr, ByVal nTimeOut As LongPtr) As LongPtr
#Else
    Private Declare Function CallNamedPipe Lib "kernel32"  Alias "CallNamedPipeA" (  ByVal lpNamedPipeName As String,  ByVal lpInBuffer As Any, ByVal nInBufferSize As Long,  ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long,  ByRef lpBytesRead As Long, ByVal nTimeOut As Long) As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As LongPtr
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
#If VBA7 Then
    Private Declare PtrSafe Function GetTempPathA Lib "kernel32"  (ByVal nBufferLength As LongPtr, ByVal lpBuffer As String) As LongPtr
#Else
    Private Declare Function GetTempPathA Lib "kernel32"  (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If
Option Explicit
' Common VBA Library - ArrayUtils
' Provides functions for handling arrays that are lacking in the built-in VBA
' language.



Private Const NORMALIZE_LBOUND = 1

' Returns a single-dimension array with lower bound 1, if given a
' one-dimensional array with any lower bound or a two-dimensional array with
' one dimension having only one element.  This function will always return a
' copy of the given array.
' Common VBA Library - ExcelUtils
' Provides useful functions for working with the Excel object model.




Public Enum Corner
    cnrTopLeft
    cnrTopRight
    cnrBottomLeft
    cnrBottomRight
End Enum

Public Enum OverwriteAction
    oaPrompt = 1
    oaOverwrite = 2
    oaSkip = 3
    oaError = 4
    oaCreateDirectory = 8
End Enum

' Determines whether a given workbook has been opened.  Pass this function
' a filename only, not a full path.
' Common VBA Library - ExportImportSheets
' Provides functions for moving sheets of one workbook to another workbook,
' breaking links along the way.  Useful for sending the results of a workbook
' to others, or for saving key data items to another workbook.


Const COL_FOLDER = 1
Const COL_FILENAME = 2
Const COL_SHEETNAME = 3
Const COL_NEWSHEETNAME = 4

' Imports Excel sheets from one or more workbooks into the given workbook,
' breaking any links and converting them to values.
' @param sheetsSpec: A two-dimensional array that describes the sheets to
' import into the workbook.  It should have one or more indices in its first
' dimension, and indices 1-3 (or 1-4) in its second dimension as follows:
'  - Index 1 (Folder) is the folder that this row's workbook appears in.
'  - Index 2 (Filename) is the filename of this row's workbook.
'  - Index 3 (SheetName) is the sheet to extract from this row's workbook.
'  - Index 4 (NewSheetName) is optional.  If present, the sheet described
'    by this row will be renamed to the NewSheetName after extraction.
' @param wb: The workbook that will receive the imported sheets (defaults to
' the workbook that contains this code).
' Common VBA Library - FileUtils
' Provides useful functions for working with filenames and paths.



' Determines whether a file with the given name exists.
' @param findFolders: If true, the function will return true if a folder
' with the given name exists.
' Common VBA Library - FormulaFunctions
' Provides functions that are useful in Excel formulas.


' Retrieves the given element of an array.
' Common VBA Library - MathFunctions
' Provides useful mathematical functions.


' Returns the lesser of its two arguments.
' Common VBA Library - StringUtils
' Provides useful functions for manipulating strings.


' Determines whether a string starts with a given prefix.
' Common VBA Library - VBAUtils
' Provides useful functions for manipulating the VBA project object model.
' @reference Microsoft Visual Basic for Applications Extensibility 5.3
' (C:\Program Files\Common Files\Microsoft Shared\vba\VBA6\VBE6.DLL)


' Determines whether a VBA code module with a given name exists.
' @param wb: The workbook to check for the given module name (defaults to the
' active workbook).
Public Function NormalizeArray(arr As Variant) As Variant
    If ArrayLen(arr) = 0 Then
        NormalizeArray = Array()
        Exit Function
    End If
    
    Dim arr2() As Variant
    
    Dim nItems As Long
    Dim i As Long
    
    Select Case ArrayRank(arr)
        Case 1
            If LBound(arr) = NORMALIZE_LBOUND Then
                NormalizeArray = arr
            Else
                nItems = ArrayLen(arr)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
            End If
            
        Case 2
            If LBound(arr, 1) = UBound(arr, 1) Then
                
                ' Copy values from array's second dimension
                nItems = ArrayLen(arr, 2)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(LBound(arr, 1), _
                        i + LBound(arr, 2) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
                
            ElseIf LBound(arr, 2) = UBound(arr, 2) Then
                
                ' Copy values from array's first dimension
                nItems = ArrayLen(arr, 1)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr, 1) - NORMALIZE_LBOUND, _
                        LBound(arr, 2))
                Next
                NormalizeArray = arr2
                
            Else
                Err.Raise 32000, Description:= _
                    "Can only normalize a 2-dimensional array if one of " _
                        & "the dimensions contains only one element."
            End If
            
        Case Else
            Err.Raise 32000, Description:= _
                "Can only normalize 1- and 2-dimensional arrays."
    End Select
End Function

' Returns the rank (number of dimensions) of an array.
' From http://www.devx.com/vb2themax/Tip/18265 .
Public Function ArrayRank(arr As Variant) As Integer
    Dim ptr As Long
    Dim vType As Integer
    Const VT_BYREF = &H4000&
    
    ' get the real VarType of the argument
    ' this is similar to VarType(), but returns also the VT_BYREF bit
    CopyMemory vType, arr, 2
    
    ' exit if not an array
    If (vType And vbArray) = 0 Then Exit Function
    
    ' get the address of the SAFEARRAY descriptor
    ' this is stored in the second half of the
    ' Variant parameter that has received the array
    CopyMemory ptr, ByVal VarPtr(arr) + 8, 4
    
    ' see whether the routine was passed a Variant
    ' that contains an array, rather than directly an array
    ' in the former case ptr already points to the SA structure.
    ' Thanks to Monte Hansen for this fix
    If (vType And VT_BYREF) Then
        ' ptr is a pointer to a pointer
        CopyMemory ptr, ByVal ptr, 4
    End If
    
    ' get the address of the SAFEARRAY structure
    ' this is stored in the descriptor
    ' get the first word of the SAFEARRAY structure
    ' which holds the number of dimensions
    ' ...but first check that saAddr is non-zero, otherwise
    ' this routine bombs when the array is uninitialized
    ' (Thanks to VB2TheMax aficionado Thomas Eyde for
    ' suggesting this edit to the original routine.)
    If ptr Then
        CopyMemory ArrayRank, ByVal ptr, 2
    End If
End Function

' Returns the number of elements in an array for a given dimension.
Public Function ArrayLen(arr As Variant, _
    Optional dimNum As Integer = 1) As Long
    
    If IsEmpty(arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function

' Sorts a section of an array in place.  Code from:
' http://stackoverflow.com/questions/152319/vba-array-sort-function
Private Sub QuickSort(vArray() As Variant, inLow As Long, inHi As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
        
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
        
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

' Sorts the given single-dimension array in place.
Public Sub SortArrayInPlace(arr() As Variant)
    QuickSort arr, LBound(arr), UBound(arr)
End Sub

' Returns a sorted copy of the given array.
Public Function SortArray(arr() As Variant) As Variant()
    If ArrayLen(arr) = 0 Then
        SortArray = Array()
    Else
        Dim arr2() As Variant
        arr2 = arr
        SortArrayInPlace arr2
        SortArray = arr2
    End If
End Function

' Returns an array containing each unique item in the given array only once.
Public Function GetUniqueItems(arr() As Variant) As Variant()
    If ArrayLen(arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim arrSorted() As Variant
        arrSorted = SortArray(arr)
        
        Dim uniqueItemsList As New zVLibList
        uniqueItemsList.Add arrSorted(LBound(arrSorted))
        
        Dim i As Long
        For i = LBound(arrSorted) + 1 To UBound(arrSorted)
            If arrSorted(i) <> arrSorted(i - 1) Then
                uniqueItemsList.Add arrSorted(i)
            End If
        Next
        
        GetUniqueItems = uniqueItemsList.Items
    End If
End Function

' Returns the subset of the given one- or two-dimensional array specified by
' the given bounds.  The returned array will have lower bounds of 1.
' @param arr: The array to process.
' @param r1: The index of the first element to be extracted from the first
' dimension of the array.  If not given, defaults to the lower bound of the
' first dimension.
' @param r2: The index of the last element to be extracted from the first
' dimension of the array.  If not given, defaults to the upper bound of the
' first dimension.
' @param c1: The index of the first element to be extracted from the second
' dimension of the array.  If not given, defaults to the lower bound of the
' second dimension.
' @param c2: The index of the last element to be extracted from the second
' dimension of the array.  If not given, defaults to the upper bound of the
' second dimension.
Public Function ArraySubset(arr() As Variant, _
    Optional r1 As Long = -1, Optional r2 As Long = -1, _
    Optional c1 As Long = -1, Optional c2 As Long = -1) As Variant()
    
    Dim arr2() As Variant
    Dim i As Long, j As Long
    
    If r1 < 0 Then r1 = LBound(arr, 1)
    If r2 < 0 Then r2 = UBound(arr, 1)
    
    Select Case ArrayRank(arr)
        Case 1
            If c1 >= 0 Or c2 >= 0 Then
                Err.Raise 32000, Description:= _
                    "Too many array dimensions passed to ArraySubset."
            End If
            ReDim arr2( _
                NORMALIZE_LBOUND To NORMALIZE_LBOUND + r2 - r1)
            For i = r1 To r2
                arr2( _
                    i - r1 + NORMALIZE_LBOUND) = arr(i)
            Next
            
        Case 2
            If c1 < 0 Then c1 = LBound(arr, 2)
            If c2 < 0 Then c2 = UBound(arr, 2)
            ReDim arr2( _
                NORMALIZE_LBOUND To NORMALIZE_LBOUND + r2 - r1, _
                NORMALIZE_LBOUND To NORMALIZE_LBOUND + c2 - c1)
            For i = r1 To r2
            For j = c1 To c2
                arr2( _
                    i - r1 + NORMALIZE_LBOUND, _
                    j - c1 + NORMALIZE_LBOUND) = arr(i, j)
            Next j, i
            
        Case Else
            Err.Raise 32000, Description:= _
                "Can only take a subset of a 1- or 2-dimensional array."
    End Select
    
    ArraySubset = arr2
End Function

' Returns the index of the given value in the given array, or one less than
' the lower bound of the array if the value is not found in the array.
' @param arr: The array to search through.
' @param val: The value to search for.
Public Function ArrayIndexOf(arr As Variant, val As Variant) As Long
    ArrayIndexOf = LBound(arr) - 1
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            ArrayIndexOf = i
            Exit Function
        End If
    Next
End Function

' Returns whether the given array contains the given value.
' @param arr: The array to search through.
' @param val: The value to search for.
Public Function ArrayContains(arr As Variant, val As Variant) As Boolean
    ArrayContains = (ArrayIndexOf(arr, val) >= LBound(arr))
End Function

Public Function IsWorkbookOpen(wbFilename As String) As Boolean
    Dim w As Workbook
    
    On Error GoTo notOpen
    Set w = Workbooks(wbFilename)
    IsWorkbookOpen = True
    Exit Function
    
notOpen:
    IsWorkbookOpen = False
End Function

' Determines whether a sheet with the given name exists.
' @param wb: The workbook to check for the given sheet name (defaults to the
' active workbook).
Public Function SheetExists(sheetName As String, Optional wb As Workbook) _
    As Boolean
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Object ' Not Worksheet because it could also be a chart
    
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function
    
notFound:
    SheetExists = False
End Function

' Determines whether a chart with the given name exists.
' @param chartName: The name of the chart to check for.
' @param sheetName: The name of the worksheet that contains the given chart
' (optional; the default is to search all worksheets).
' @param wb: The workbook to check for the given chart name (defaults to the
' active workbook.
Public Function ChartExists(chartName As String, _
    Optional sheetName As String = "", Optional wb As Workbook) As Boolean
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim s As Worksheet
    Dim c As ChartObject
    
    ChartExists = False
    
    If sheetName = "" Then
        For Each s In wb.Sheets
            If ChartExists(chartName, s.Name, wb) Then
                ChartExists = True
                Exit Function
            End If
        Next
    Else
        Set s = wb.Sheets(sheetName)
        On Error GoTo notFound
        Set c = s.ChartObjects(chartName)
        ChartExists = True
notFound:
    End If
End Function

' Deletes the sheet with the given name, without prompting for confirmation.
' @param wb: The workbook to check for the given sheet name (defaults to the
' active workbook).
Public Sub DeleteSheetByName(sheetName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then DeleteSheetOrSheets wb.Sheets(sheetName)
End Sub

' Deletes the given worksheet, without prompting for confirmation.
Public Sub DeleteSheet(s As Worksheet)
    DeleteSheetOrSheets s
End Sub

' Deletes all sheets in the given Sheets object, without prompting for
' confirmation.
Public Sub DeleteSheets(s As Sheets)
    DeleteSheetOrSheets s
End Sub

Private Sub DeleteSheetOrSheets(s As Object)
    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    s.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
End Sub

' Returns the actual used range from a sheet.
' @param fromTopLeft: If True, returns the used range starting from cell A1,
' which is different from the way Excel's UsedRange property behaves if the
' sheet does not use any cells in the top row(s) and/or leftmost column(s).
Public Function GetRealUsedRange(s As Worksheet, _
    Optional fromTopLeft As Boolean = True) As Range
    
    If fromTopLeft Then
        Set GetRealUsedRange = s.Range( _
            s.Cells(1, 1), _
            s.Cells( _
                s.UsedRange.Rows.Count + s.UsedRange.Row - 1, _
                s.UsedRange.Columns.Count + s.UsedRange.Column - 1))
    Else
        Set GetRealUsedRange = s.UsedRange
    End If
End Function

' Sets the value of the given range if it is different than the proposed value.
' Returns whether the value of the range was changed.
Public Function SetValueIfNeeded(rng As Range, val As Variant) As Boolean
    If rng.Value = val Then
        SetValueIfNeeded = False
    Else
        rng.Value = val
        SetValueIfNeeded = True
    End If
End Function

' Converts an integer column number to an Excel column string.
Public Function ExcelCol(c As Integer) As String
    ExcelCol = ExcelCol_ZeroBased(c - 1)
End Function

Private Function ExcelCol_ZeroBased(c As Integer) As String
    Dim c2 As Integer
    c2 = c \ 26
    If c2 = 0 Then
        ExcelCol_ZeroBased = Chr(65 + c)
    Else
        ExcelCol_ZeroBased = ExcelCol(c2) & Chr(65 + (c Mod 26))
    End If
End Function

' Converts an Excel column string to an integer column number.
Public Function ExcelColNum(c As String) As Integer
    ExcelColNum = 0
    Dim i As Integer
    For i = 1 To Len(c)
        ExcelColNum = (ExcelColNum + Asc(Mid(c, i, 1)) - 64)
        If i < Len(c) Then ExcelColNum = ExcelColNum * 26
    Next
End Function

' Builds an Excel cell reference.
Public Function CellReference(ByVal r As Long, ByVal c As Integer, _
    Optional sheet As String = "", Optional absoluteRow As Boolean = False, _
    Optional absoluteCol As Boolean = False) As String
    
    Dim ref As String
    ref = IIf(absoluteCol, "$", "") & ExcelCol(c) _
        & IIf(absoluteRow, "$", "") & r
    
    If sheet = "" Then
        CellReference = ref
    Else
        CellReference = "'" & Replace(sheet, "'", "''") & "'!" & ref
    End If
End Function

' Returns a string describing the type of an Excel error value
' ("#DIV/0!", "#N/A", etc.)
Public Function ExcelErrorType(e As Variant) As String
    If IsError(e) Then
        Select Case e
            Case CVErr(xlErrDiv0)
                ExcelErrorType = "#DIV/0!"
            Case CVErr(xlErrNA)
                ExcelErrorType = "#N/A"
            Case CVErr(xlErrName)
                ExcelErrorType = "#NAME?"
            Case CVErr(xlErrNull)
                ExcelErrorType = "#NULL!"
            Case CVErr(xlErrNum)
                ExcelErrorType = "#NUM!"
            Case CVErr(xlErrRef)
                ExcelErrorType = "#REF!"
            Case CVErr(xlErrValue)
                ExcelErrorType = "#VALUE!"
            Case Else
                ExcelErrorType = "#UNKNOWN_ERROR"
        End Select
    Else
        ExcelErrorType = "(not an error)"
    End If
End Function

' Shows a status message to update the user on the progress of a long-running
' operation, in a way that can be detected by external applications.
Public Sub ShowStatusMessage(statusMessage As String)
    ' Show the message in the status bar.
    Application.StatusBar = statusMessage
    
    ' Set the Excel window title to the updated status message.  The window
    ' title as seen by the Windows API will then be:
    ' "Status Message - WorkbookFilename.xlsm"
    ' To allow external applications to extract just the status message,
    ' put the length of the message at the beginning.
    Application.Caption = Len(statusMessage) & ":" & statusMessage
End Sub

' Shows a status message for 2-3 seconds then removes it.
Public Sub FlashStatusMessage(statusMessage As String)
    ShowStatusMessage statusMessage
    Application.OnTime Now + TimeValue("0:00:02"), "ClearStatusMessage"
End Sub

' Clears any status message that is currently being displayed by a macro.
Public Sub ClearStatusMessage()
    Application.StatusBar = False
    Application.Caption = Empty
End Sub

' Attempts to send a message to an external program that is running this macro
' and listening for messages.
Public Sub SendMessageToListener(msg As String)
    Dim bArray(0 To 0) As Byte

    #If VBA7 Then
        Dim bytesRead As LongPtr
    #Else
        Dim bytesRead As Long
    #End If

    CallNamedPipe _
        "\\.\pipe\ExcelMacroCommunicationListener." & GetCurrentProcessId, _
        msg, Len(msg), bArray(0), 1, bytesRead, 500
End Sub

' Returns the cell in the given corner of the given range.
Public Function GetCornerCell(r As Range, c As Corner) As Range
    Select Case c
        Case cnrTopLeft
            Set GetCornerCell = r.Cells(1, 1)
        Case cnrTopRight
            Set GetCornerCell = r.Cells(1, r.Columns.Count)
        Case cnrBottomLeft
            Set GetCornerCell = r.Cells(r.Rows.Count, 1)
        Case cnrBottomRight
            Set GetCornerCell = r.Cells(r.Rows.Count, r.Columns.Count)
    End Select
End Function

' Returns an array of objects representing the other Excel workbooks that the
' given workbook links to.
' @param wb: The source workbook (defaults to the active workbook).
Public Function GetAllExcelLinks(Optional wb As Workbook) As Variant
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim linkNames() As Variant
    linkNames = NormalizeArray(ActiveWorkbook.LinkSources(xlExcelLinks))
    
    If ArrayLen(linkNames) Then
        Dim linksArr() As zVLibLink
        ReDim linksArr(1 To ArrayLen(linkNames))
        Dim i As Integer
        For i = 1 To UBound(linkNames)
            Set linksArr(i) = New zVLibLink
            linksArr(i).Initialize wb, CStr(linkNames(i))
        Next
        GetAllExcelLinks = linksArr
    Else
        GetAllExcelLinks = Array()
        Exit Function
    End If
End Function

Private Function GetMatchingLinkName(linkFilename As String, _
    Optional wb As Workbook) As String
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim linkNames() As Variant
    linkNames = NormalizeArray(wb.LinkSources(xlExcelLinks))
    
    Dim i As Integer, matchingLinkName As String
    
    ' First look for a link with the exact full path given by linkFilename
    For i = 1 To UBound(linkNames)
        If LCase(linkNames(i)) = LCase(linkFilename) Then
            GetMatchingLinkName = linkNames(i)
            Exit Function
        End If
    Next
    
    ' Next look for a link with the same filename as linkFilename.  Do it in
    ' two steps because it is actually possible for Excel to link to two
    ' workbooks with the same name in different folders.  No one should ever
    ' do this, but we'll try to support retrieving such links anyway.
    For i = 1 To UBound(linkNames)
        If LCase(GetFilename(linkNames(i))) = _
            LCase(GetFilename(linkFilename)) Then
            
            GetMatchingLinkName = linkNames(i)
            Exit Function
        End If
    Next
    
    GetMatchingLinkName = ""
End Function

' Returns an object representing the link to the Excel workbook with the given
' filename.
' @param linkFilename: The path or filename of the linked Excel workbook.
' @param wb: The workbook that contains the link (defaults to the active
' workbook).
Public Function GetExcelLink(linkFilename As String, Optional wb As Workbook) _
    As zVLibLink
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim matchingLinkName As String
    matchingLinkName = GetMatchingLinkName(linkFilename, wb)
    
    If matchingLinkName = "" Then
        Err.Raise 32000, Description:= _
            "No Excel link exists with the given name ('" & linkFilename _
                & "')."
    Else
        Set GetExcelLink = New zVLibLink
        GetExcelLink.Initialize wb, matchingLinkName
    End If
End Function

' Returns whether an Excel link matching the given workbook filename exists.
' @param wb: The workbook that contains the link (defaults to the active
' workbook).
Public Function ExcelLinkExists(linkFilename As String, _
    Optional wb As Workbook) As Boolean
    
    ExcelLinkExists = (GetMatchingLinkName(linkFilename, wb) <> "")
End Function

' Refreshes all Access database connections in the given workbook.
' @param wb: The workbook to refresh (defaults to the active workbook).
Public Sub RefreshAccessConnections(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim cn As WorkbookConnection
    
    On Error GoTo err_
    Application.Calculation = xlCalculationManual
    
    Dim numConnections As Integer, i As Integer
    
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            numConnections = numConnections + 1
        End If
    Next
    
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            i = i + 1
            ShowStatusMessage "Refreshing data connection '" _
                & cn.OLEDBConnection.CommandText _
                & "' (" & i & " of " & numConnections & ")"
            cn.OLEDBConnection.BackgroundQuery = False
            cn.Refresh
       End If
    Next
    
    GoTo done_
err_:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
done_:
    ShowStatusMessage "Recalculating"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    ClearStatusMessage
End Sub

' Returns a wrapper object for the table with the given name in the given
' workbook.
' @param wb: The workbook that contains the table (defaults to the active
' workbook).
Public Function GetExcelTable(tblName As String, Optional wb As Workbook) _
    As zVLibTable
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error GoTo notFound
    
    Dim wbPrevActive As Workbook
    Set wbPrevActive = ActiveWorkbook
    wb.Activate
    
    ' We could just do Range(tblName).ListObject, but then this would allow
    ' getting a table by any of its cells or columns.  Instead, do some
    ' verification that the string we were passed is actually the name of
    ' a table.
    Dim tbl As ListObject
    Set tbl = Range(tblName).Parent.ListObjects(tblName)
    
    wbPrevActive.Activate
    
    Set GetExcelTable = New zVLibTable
    GetExcelTable.Initialize tbl
    Exit Function
    
notFound:
    On Error GoTo 0
    Err.Raise 32000, Description:= _
        "Could not find table '" & tblName & "'."
End Function

' Returns the Excel workbook format for the given file extension.
Public Function GetWorkbookFileFormat(fileExtension As String) As XlFileFormat
    Select Case LCase(Replace(fileExtension, ".", ""))
        Case "xls"
            GetWorkbookFileFormat = xlExcel8
        Case "xla"
            GetWorkbookFileFormat = xlAddIn8
        Case "xlt"
            GetWorkbookFileFormat = xlTemplate8
        Case "csv"
            GetWorkbookFileFormat = xlCSV
        Case "txt"
            GetWorkbookFileFormat = xlCurrentPlatformText
        Case "xlsx"
            GetWorkbookFileFormat = xlOpenXMLWorkbook
        Case "xlsm"
            GetWorkbookFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xlsb"
            GetWorkbookFileFormat = xlExcel12
        Case "xlam"
            GetWorkbookFileFormat = xlOpenXMLAddIn
        Case "xltx"
            GetWorkbookFileFormat = xlOpenXMLTemplate
        Case "xltm"
            GetWorkbookFileFormat = xlOpenXMLTemplateMacroEnabled
        Case Else
            Err.Raise 32000, Description:= _
                "Unrecognized Excel file extension: '" & fileExtension & "'"
    End Select
End Function

' Saves the given workbook as a different filename, with options for handling
' the case where the file already exists.  Returns True if the workbook was
' saved, or False if it was not saved.
' @param oAction: The action that will be taken if the given file exists.  This
' parameter also accepts the oaCreateDirectory flag, which means that the
' directory hierarchy of the requested filename will be created if it does not
' already exist.  If not given, defaults to oaPrompt.
Public Function SaveWorkbookAs(wb As Workbook, newFilename As String, _
    Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    If Not FolderExists(GetDirectoryName(newFilename)) Then
        If oAction And oaCreateDirectory Then
            MkDirRecursive GetDirectoryName(newFilename)
        Else
            Err.Raise 32000, Description:= _
                "The parent folder of the requested workbook filename " _
                    & "does not exist:" & vbLf & vbLf & newFilename
        End If
    End If
    
    If FileExists(newFilename) Then
        If oAction And oaOverwrite Then
            Kill newFilename
            ' Proceed to save the file
            
        ElseIf oAction And oaError Then
            Err.Raise 32000, Description:= _
                "The given filename already exists:" _
                    & vbLf & vbLf & newFilename
            
        ElseIf oAction And oaPrompt Then
            Dim r As VbMsgBoxResult
            r = MsgBox(Title:="Overwrite Excel file?", _
                Buttons:=vbYesNo + vbExclamation, _
                Prompt:="The following Excel file already exists:" _
                    & vbLf & vbLf & newFilename & vbLf & vbLf _
                    & "Do you want to overwrite it?")
            If r = vbYes Then
                Kill newFilename
                ' Proceed to save the file
            Else
                SaveWorkbookAs = False
                Exit Function
            End If
            
        ElseIf oAction And oaSkip Then
            SaveWorkbookAs = False
            Exit Function
            
        Else
            Err.Raise 32000, Description:= _
                "Bad overwrite action value passed to SaveWorkbookAs."
            
        End If
    End If
    
    ' wb.SaveCopyAs doesn't take all the fancy arguments that wb.SaveAs does,
    ' but it's the only way to save a copy of the current workbook.  This
    ' means, among other things, that it is not possible to save the workbook
    ' as a different format than the original workbook.  To work around this,
    ' call SaveCopyAs with a temporary filename first, then open the temporary
    ' file, then call SaveAs with the desired filename and options.
    
    Dim tmpFilename As String
    tmpFilename = CombinePaths(GetTempPath, Int(Rnd * 1000000) & "-" & wb.Name)
    wb.SaveCopyAs tmpFilename
    
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Open(tmpFilename, UpdateLinks:=False, ReadOnly:=True)
    wbTmp.SaveAs filename:=newFilename, _
        FileFormat:=GetWorkbookFileFormat(GetFileExtension(newFilename)), _
        ReadOnlyRecommended:=openReadOnly
    wbTmp.Close SaveChanges:=False
    
    Kill tmpFilename
    
    SaveWorkbookAs = True
End Function

' Saves the current workbook as a different filename, with options for handling
' the case where the file already exists.  Returns True if the workbook was
' saved, or False if it was not saved.
' @param oAction: The action that will be taken if the given file exists.  This
' parameter also accepts the oaCreateDirectory flag, which means that the
' directory hierarchy of the requested filename will be created if it does not
' already exist.  If not given, defaults to oaPrompt.
' @param openReadOnly: True or False to determine whether the created workbook
' will prompt users to open it as read-only (defaults to False).
Public Function SaveActiveWorkbookAs(newFilename As String, _
    Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    SaveActiveWorkbookAs = SaveWorkbookAs(ActiveWorkbook, newFilename, _
        oAction, openReadOnly)
End Function

Public Sub ImportExcelSheets(sheetsSpec() As Variant, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    CopyExcelSheets wb, sheetsSpec, False
End Sub

' Exports Excel sheets from one or more workbooks into a new workbook,
' breaking any links and converting them to values.  Returns True if the
' workbook was saved, or False if it was not saved.
' @param sheetsSpec: A two-dimensional array that describes the sheets to
' export to the new workbook.  This array has the same structure as in
' ImportExcelSheets, with the added feature that if the folder is blank and
' the filename is the string ThisWorkbook then the sheet to be exported will
' come from the current workbook.
' @param wb: The workbook that will receive the imported sheets (defaults to
' the active workbook).
Public Function ExportExcelSheets(sheetsSpec() As Variant, _
    wbFilename As String, Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    While wb.Sheets.Count > 1
        DeleteSheet wb.Sheets(2)
    Wend
    
    CopyExcelSheets wb, sheetsSpec, True
    
    DeleteSheet wb.Sheets(1)
    
    ExportExcelSheets = SaveWorkbookAs(wb, wbFilename, oAction, openReadOnly)
    
    wb.Close SaveChanges:=False
End Function

Private Sub CopyExcelSheets(wb As Workbook, sheetsSpec() As Variant, _
    allowThisWorkbook As Boolean)
    
    Dim prevActiveSheet As Worksheet
    Set prevActiveSheet = wb.ActiveSheet
    
    Dim i1 As Long, i2 As Long
    i1 = LBound(sheetsSpec, 1)
    i2 = UBound(sheetsSpec, 1)
    
    ' The workbooks which contain the sheets we're interested in
    Dim wbFilenames() As String
    ReDim wbFilenames(i1 To i2)
    
    ' The names of the sheets we're interested in
    Dim sheetNames() As String
    ReDim sheetNames(i1 To i2)
    
    ' The new names of the sheets we're interested in
    Dim newSheetNames() As String
    ReDim newSheetNames(i1 To i2)
    
    ' The desired position of each sheet (this array stores a sheet name that
    ' the sheet will be placed immediately after, or the empty string if the
    ' sheet should be placed at the beginning of the workbook)
    Dim sheetPositions() As String
    ReDim sheetPositions(i1 To i2)
    
    ' The order in which sheets need to be moved when they are rearranged.  To
    ' see why this is necessary, imagine that a workbook contains sheets A, B,
    ' C, and D, but the program obtains these sheets in the order A, C, B, D.
    ' When rearranging sheets, A would be moved to its position (correctly),
    ' then C would be moved to its position after B, but since B was not in the
    ' desired position, then C would not be moved to its desired position
    ' either.  To solve this, store the order of the existing sheets in the
    ' workbook, and move the new sheets in that order.
    Dim sheetMoveOrder() As Long
    ReDim sheetMoveOrder(i1 To i2)
    ' Supporting variables for sheetMoveOrder.
    Dim sheetIndex As Long, sheetMoveOrderIndex As Long
    
    ' The list of Excel links to other workbooks that could not be broken.
    Dim linksFailedToBreak As New zVLibList
    
    Dim i As Long
    For i = i1 To i2
        Dim thisFolderName As String, thisFilename As String
        thisFolderName = CStr(sheetsSpec(i, COL_FOLDER))
        thisFilename = CStr(sheetsSpec(i, COL_FILENAME))
        If thisFolderName = "" Or thisFilename = "" Then
            wbFilenames(i) = thisFolderName & thisFilename
        Else
            wbFilenames(i) = NormalizePath(CombinePaths( _
                thisFolderName, thisFilename))
        End If
        
        If LCase(wbFilenames(i)) = "thisworkbook" Then
            If allowThisWorkbook Then
                wbFilenames(i) = "ThisWorkbook"
            Else
                Err.Raise 32000, Description:= _
                    "Cannot import Excel sheets from ThisWorkbook."
            End If
        End If
        
        sheetNames(i) = sheetsSpec(i, COL_SHEETNAME)
        
        If COL_NEWSHEETNAME <= UBound(sheetsSpec, 2) Then
            newSheetNames(i) = sheetsSpec(i, COL_NEWSHEETNAME)
        Else
            newSheetNames(i) = ""
        End If
        If newSheetNames(i) = "" Then
            newSheetNames(i) = sheetNames(i)
        End If
        
        If SheetExists(newSheetNames(i), wb) Then
            sheetIndex = wb.Sheets(newSheetNames(i)).Index
            If sheetIndex = 1 Then
                sheetPositions(i) = ""
            Else
                sheetPositions(i) = wb.Sheets(sheetIndex - 1).Name
            End If
        ElseIf i = i1 Then
            sheetPositions(i) = wb.Sheets(wb.Sheets.Count).Name
        Else
            sheetPositions(i) = newSheetNames(i - 1)
        End If
        
        sheetMoveOrder(i) = i1 - 1
    Next
    
    ' Determine the order in which we need to rearrange sheets.  Start by
    ' looping over all of the workbook's current sheets, and checking if they
    ' are sheets that will be replaced during this run.  If so, then rearrange
    ' them in that order.
    sheetMoveOrderIndex = i1
    For sheetIndex = 1 To wb.Sheets.Count
        i = ArrayIndexOf(newSheetNames, wb.Sheets(sheetIndex).Name)
        If i >= i1 Then
            sheetMoveOrder(sheetMoveOrderIndex) = i
            sheetMoveOrderIndex = sheetMoveOrderIndex + 1
        End If
    Next
    ' Now, add any sheets that will be added to the workbook, but do not exist
    ' yet.  Ensure that these sheets are arranged according to the order in
    ' the specification passed to this function.
    For i = i1 To i2
        If Not ArrayContains(sheetMoveOrder, i) Then
            sheetMoveOrder(sheetMoveOrderIndex) = i
            sheetMoveOrderIndex = sheetMoveOrderIndex + 1
        End If
    Next
    ' Sanity check
    'If sheetMoveOrderIndex <> i2 + 1 Then Stop
    
    Dim currentFilename As String
    Dim currentWb As Workbook
    Dim filesProcessed As New zVLibList
    Dim sheetsToCopy As New zVLibList
    Dim oldLinkNames As New zVLibList
    Dim newLinkNames As New zVLibList
    
    Do
        currentFilename = ""
        sheetsToCopy.Clear
        
        For i = i1 To i2
            If currentFilename = "" Then
                If Not filesProcessed.Contains(LCase(wbFilenames(i))) Then
                    currentFilename = wbFilenames(i)
                    filesProcessed.Add LCase(wbFilenames(i))
                    Set currentWb = Nothing
                End If
            End If
            
            If LCase(currentFilename) = LCase(wbFilenames(i)) Then
                If currentWb Is Nothing Then
                    If currentFilename = "ThisWorkbook" Then
                        Set currentWb = ThisWorkbook
                    Else
                        ShowStatusMessage "Opening workbook: " _
                            & GetFilename(currentFilename)
                        If IsWorkbookOpen(GetFilename(currentFilename)) Then
                            Set currentWb = Workbooks( _
                                GetFilename(currentFilename))
                        Else
                            Set currentWb = Workbooks.Open(wbFilenames(i), _
                                ReadOnly:=True, _
                                UpdateLinks:=False)
                        End If
                        ClearStatusMessage
                    End If
                End If
                
                If SheetExists(newSheetNames(i), wb) Then
                    ShowStatusMessage "Deleting sheet: " & newSheetNames(i)
                    DeleteSheetByName newSheetNames(i), wb
                    ClearStatusMessage
                End If
                
                ' Instead of copying sheets one at a time, save a list of the
                ' sheet names we need to copy and do them all at once.  This
                ' way is much faster.
                sheetsToCopy.Add sheetNames(i)
            End If
        Next
        
        If Not currentWb Is Nothing Then
            Dim oldSheetCount As Long
            oldSheetCount = wb.Sheets.Count
            
            ' Store the list of linked files in the workbook before copying
            ' sheets over, because copying a sheet can add more than one link.
            oldLinkNames.Clear
            On Error Resume Next ' wb.LinkSources returns Empty if no links
            oldLinkNames.AddRange wb.LinkSources(xlExcelLinks)
            On Error GoTo 0
            
            ShowStatusMessage "Copying sheets from workbook: " _
                & currentWb.Name
            currentWb.Sheets(sheetsToCopy.Items).Copy _
                After:=wb.Sheets(wb.Sheets.Count)
            
            ' Unhide any sheets that were hidden when copied over.
            For i = oldSheetCount + 1 To wb.Sheets.Count
                wb.Sheets(i).Visible = xlSheetVisible
            Next
            
            For i = i1 To i2
                If LCase(currentFilename) = LCase(wbFilenames(i)) _
                    And sheetNames(i) <> newSheetNames(i) Then
                    
                    ShowStatusMessage "Renaming sheet: " & newSheetNames(i)
                    wb.Sheets(sheetNames(i)).Name = newSheetNames(i)
                End If
            Next
            
            ' Get the list of links again, and remove any that didn't exist
            ' before, as well as any link to the workbook that contains the
            ' sheet(s) we're currently copying.
            newLinkNames.Clear
            newLinkNames.AddRange wb.LinkSources(xlExcelLinks)
            If ExcelLinkExists(currentWb.Name, wb) Then
                newLinkNames.AddOnce currentWb.FullName
            End If
            
            Dim linkName_ As Variant, linkName As String
            For Each linkName_ In newLinkNames.Items
                linkName = linkName_
                ' Always try to remove the link to the current workbook, even
                ' if it already existed.
                If LCase(GetFilename(linkName)) = LCase(currentWb.Name) _
                    Or Not oldLinkNames.Contains(linkName) Then
                    
                    ShowStatusMessage "Breaking link to workbook: " _
                        & GetFilename(linkName)
                    Dim currentWbLink As zVLibLink
                    Set currentWbLink = GetExcelLink(linkName, wb)
                    If Not currentWbLink.Break(False) Then
                        linksFailedToBreak.Add GetFilename(linkName)
                    End If
                End If
            Next
            
            If currentFilename <> "ThisWorkbook" Then
                ShowStatusMessage "Closing workbook: " & currentWb.Name
                currentWb.Close SaveChanges:=False
            End If
            Set currentWb = Nothing
            
            ClearStatusMessage
        End If
    Loop While currentFilename <> ""
    
    ShowStatusMessage "Rearranging sheets"
    For i = i1 To i2
        If sheetPositions(sheetMoveOrder(i)) = "" Then
            wb.Sheets(newSheetNames(sheetMoveOrder(i))).Move _
                Before:=wb.Sheets(1)
        ElseIf SheetExists(sheetPositions(sheetMoveOrder(i)), wb) Then
            wb.Sheets(newSheetNames(sheetMoveOrder(i))).Move _
                After:=wb.Sheets(sheetPositions(sheetMoveOrder(i)))
        End If
    Next
    ClearStatusMessage
    
    If linksFailedToBreak.Count Then
        MsgBox Prompt:="Failed to break links to one or more workbooks:" _
                & vbLf & vbLf & Join(linksFailedToBreak.Items, vbLf), _
            Title:="Excel link failure", _
            Buttons:=vbOKOnly + vbExclamation
    End If
    
    prevActiveSheet.Activate
End Sub

Public Function FileExists(ByVal testFilename As String, _
    Optional findFolders As Boolean = False) As Boolean
    
    ' Include read-only files, hidden files, system files.
    Dim attrs As Long
    attrs = (vbReadOnly Or vbHidden Or vbSystem)
    
    If findFolders Then
        attrs = (attrs Or vbDirectory) ' Include folders as well.
    End If
    
    'If Dir() returns something, the file exists.
    FileExists = False
    On Error Resume Next
    FileExists = (Dir(TrimTrailingChars(testFilename, "/\"), attrs) <> "")
End Function

' Determines whether a folder with the given name exists.
Public Function FolderExists(folderName As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderName) And vbDirectory) = vbDirectory)
End Function

' Creates the given directory, including any missing parent folders.
Public Sub MkDirRecursive(folderName As String)
    MkDirRecursiveInternal folderName, folderName
End Sub

Private Sub MkDirRecursiveInternal(folderName As String, _
    originalFolderName As String)
    
    If folderName = "" Then
        ' Too many recursive calls to this function (GetDirectoryName will
        ' eventually return an empty string)
        Err.Raise 32000, _
            Description:="Failed to create folder: " & originalFolderName
    End If
    
    Dim parentFolderName As String
    parentFolderName = GetDirectoryName(folderName)
    If Not FolderExists(parentFolderName) Then
        MkDirRecursiveInternal parentFolderName, originalFolderName
    End If
    
    If Not FolderExists(folderName) Then
        MkDir folderName
    End If
End Sub

' Merges two path components into a single path.
Public Function CombinePaths(p1 As String, p2 As String) As String
    CombinePaths = _
        TrimTrailingChars(p1, "/\") & "\" & _
        TrimLeadingChars(p2, "/\")
End Function

' Fixes slashes within a path:
'  - Converts all forward slashes to backslashes
'  - Removes multiple consecutive slashes (except for UNC paths)
'  - Removes any trailing slashes
Public Function NormalizePath(ByVal p As String) As String
    Dim isUNC As Boolean
    isUNC = StartsWith(p, "\\")
    p = Replace(p, "/", "\")
    While InStr(p, "\\") > 0
        p = Replace(p, "\\", "\")
    Wend
    If isUNC Then p = "\" & p
    NormalizePath = TrimTrailingChars(p, "\")
End Function

' Returns the folder name of a path (removes the last component
' of the path).
Public Function GetDirectoryName(ByVal p As String) As String
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    If i = 0 Then
        GetDirectoryName = ""
    Else
        GetDirectoryName = Left(p, i - 1)
    End If
End Function

' Returns the filename of a path (the last component of the path).
Public Function GetFilename(ByVal p As String) As String
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    GetFilename = Mid(p, i + 1)
End Function

' Returns the extension of a filename (including the dot).
Public Function GetFileExtension(ByVal p As String) As String
    Dim i As Integer
    i = InStrRev(p, ".")
    If i > 0 Then
        GetFileExtension = Mid(p, i)
    Else
        GetFileExtension = ""
    End If
End Function

Private Function ListFiles_Internal(filePattern As String, attrs As Long) _
    As Variant()
    
    Dim filesList As New zVLibList
    Dim folderName As String
    
    If FolderExists(filePattern) Then
        filePattern = NormalizePath(filePattern) & "\"
        folderName = filePattern
    Else
        folderName = GetDirectoryName(filePattern) & "\"
    End If
    
    Dim currFilename As String
    currFilename = Dir(filePattern, attrs)
    
    While currFilename <> ""
        If (attrs And vbDirectory) = vbDirectory Then
            If FolderExists(folderName & currFilename) _
                And currFilename <> "." And currFilename <> ".." Then
                
                filesList.Add folderName & currFilename
            End If
        Else
            filesList.Add folderName & currFilename
        End If
        currFilename = Dir
    Wend
    
    ListFiles_Internal = filesList.Items
End Function

' Lists all files matching the given pattern.
' @param filePattern: A directory name, or a path with wildcards:
'  - C:\Path\to\Folder
'  - C:\Path\to\Folder\ExcelFiles.xl*
Public Function ListFiles(filePattern As String) As Variant()
    ListFiles = ListFiles_Internal(filePattern, _
        vbReadOnly Or vbHidden Or vbSystem)
End Function

' Lists all folders matching the given pattern.
' @param folderPattern: A directory name, or a path with wildcards:
'  - C:\Path\to\Folder
'  - C:\Path\to\Folder\OtherFolder_*
Public Function ListFolders(folderPattern As String) As Variant()
    ListFolders = ListFiles_Internal(folderPattern, _
        vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
End Function

' Returns the path to a folder that can be used to store temporary
' files.
Public Function GetTempPath() As String
    Const MAX_PATH = 256
    
    Dim folderName As String

    #If VBA7 Then
        Dim ret As LongPtr
    #Else
        Dim ret As Long
    #End If

    
    folderName = String(MAX_PATH, 0)
    ret = GetTempPathA(MAX_PATH, folderName)
    
    If ret <> 0 Then
        GetTempPath = Left(folderName, InStr(folderName, Chr(0)) - 1)
    Else
        Err.Raise 32000, Description:= _
            "Error getting temporary folder."
    End If
End Function

Public Function ArrayElement(arr As Variant, i1 As Variant, _
    Optional i2 As Variant, Optional i3 As Variant, _
    Optional i4 As Variant, Optional i5 As Variant) As Variant
    
    If IsMissing(i2) Then
        If IsObject(arr(i1)) Then
            Set ArrayElement = arr(i1)
        Else
            ArrayElement = arr(i1)
        End If
    ElseIf IsMissing(i3) Then
        If IsObject(arr(i1, i2)) Then
            Set ArrayElement = arr(i1, i2)
        Else
            ArrayElement = arr(i1, i2)
        End If
    ElseIf IsMissing(i4) Then
        If IsObject(arr(i1, i2, i3)) Then
            Set ArrayElement = arr(i1, i2, i3)
        Else
            ArrayElement = arr(i1, i2, i3)
        End If
    ElseIf IsMissing(i5) Then
        If IsObject(arr(i1, i2, i3, i4)) Then
            Set ArrayElement = arr(i1, i2, i3, i4)
        Else
            ArrayElement = arr(i1, i2, i3, i4)
        End If
    Else
        If IsObject(arr(i1, i2, i3, i4, i5)) Then
            Set ArrayElement = arr(i1, i2, i3, i4, i5)
        Else
            ArrayElement = arr(i1, i2, i3, i4, i5)
        End If
    End If
End Function

' Splits a string into an array, optionally limiting the number
' of items in the returned array.
Public Function StringSplit(s As String, delim As String, _
    Optional limit As Long = -1) As String()
    
    StringSplit = Split(s, delim, limit)
End Function

' Joins an array into a string by inserting the given delimiter in between
' items.
Public Function StringJoin(arr() As Variant, delim As String) As String
    StringJoin = Join(arr, delim)
End Function

' Returns a newline (vbLf) character for use in formulas.
Public Function NewLine() As String
    NewLine = vbLf
End Function

' Returns an array suitable for using in an array formula.  When this
' function is called from an array formula, it will detect whether or
' not the array should be transposed to fit into the range.
Public Function GetArrayForFormula(arr As Variant) As Variant
    If IsObject(Application.Caller) Then
        Dim len1 As Long, len2 As Long
        Select Case ArrayRank(arr)
            Case 0
                GetArrayForFormula = Empty
                Exit Function
            Case 1
                len1 = ArrayLen(arr)
                len2 = 1
            Case 2
                len1 = ArrayLen(arr)
                len2 = ArrayLen(arr, 2)
            Case Else
                Err.Raise 32000, Description:= _
                    "Invalid number of dimensions (" & ArrayRank(arr) _
                        & "; expected 1 or 2)."
        End Select
        
        If Application.Caller.Rows.Count > Application.Caller.Columns.Count _
            And len1 > len2 Then
            
            GetArrayForFormula = WorksheetFunction.Transpose(arr)
        Else
            GetArrayForFormula = arr
        End If
    Else
        GetArrayForFormula = arr
    End If
End Function

' Converts a range to a normalized array.
Public Function RangeToArray(r As Range) As Variant()
    If r.Cells.Count = 1 Then
        RangeToArray = Array(r.Value)
    ElseIf r.Rows.Count = 1 Or r.Columns.Count = 1 Then
        RangeToArray = NormalizeArray(r.Value)
    Else
        RangeToArray = r.Value
    End If
End Function

' Returns the width of a column on a sheet.  If the column number is
' not given and this function is used in a formula, then it returns
' the column width of the cell containing the formula.
Public Function ColumnWidth(Optional c As Integer = 0) As Variant
    Application.Volatile
    Dim s As Worksheet
    If IsObject(Application.Caller) Then
        Set s = Application.Caller.Worksheet
    Else
        Set s = ActiveSheet
    End If
    If c <= 0 And IsObject(Application.Caller) Then
        c = Application.Caller.Column
    End If
    ColumnWidth = s.Columns(c).Width
End Function

' Returns the height of a row on a sheet.  If the row number is
' not given and this function is used in a formula, then it returns
' the row height of the cell containing the formula.
Public Function RowHeight(Optional r As Integer = 0) As Variant
    Application.Volatile
    Dim s As Worksheet
    If IsObject(Application.Caller) Then
        Set s = Application.Caller.Worksheet
    Else
        Set s = ActiveSheet
    End If
    If r <= 0 And IsObject(Application.Caller) Then
        r = Application.Caller.Row
    End If
    RowHeight = s.Rows(r).Height
End Function

' Returns the formula of the given cell or range, optionally in R1C1 style.
Public Function GetFormula(r As Range, Optional r1c1 As Boolean = False) _
    As Variant
    
    If r1c1 Then
        GetFormula = r.FormulaR1C1
    Else
        GetFormula = r.Formula
    End If
End Function

Public Function Min(a As Double, b As Double) As Double
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

' Returns the greater of its two arguments.
Public Function Max(a As Double, b As Double) As Double
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

' Returns its argument truncated (rounded down) to the given significance or
' the given number of decimal places.
' @param significance: The significance, or step size, of the function.  For
' example, a step size of 0.2 will ensure that the number returned is a
' multiple of 0.2.
' @param places: The number of decimal places to keep.
Public Function Floor(num As Double, _
    Optional significance As Double = 1, _
    Optional places As Integer = 0) As Double
    
    ValidateFloorCeilingParams significance, places
    Floor = Int(num / significance) * significance
End Function

' Returns its argument rounded up to the given significance or the given number
' of decimal places.
' @param significance: The significance, or step size, of the function.  For
' example, a step size of 0.2 will ensure that the number returned is a
' multiple of 0.2.
' @param places: The number of decimal places to keep.
Public Function Ceiling(num As Double, _
    Optional significance As Double = 1, _
    Optional places As Integer = 0) As Double
    
    ValidateFloorCeilingParams significance, places
    Ceiling = Floor(num, significance)
    If num <> Ceiling Then Ceiling = Ceiling + significance
End Function

Private Sub ValidateFloorCeilingParams( _
    ByRef significance As Double, _
    ByRef places As Integer)
    
    If places <> 0 Then
        If significance <> 1 Then
            Err.Raise 32000, Description:= _
                "Pass either a number of decimal places or a significance " _
                & "to Floor() or Ceiling(), not both."
        Else
            significance = 10 ^ -places
        End If
    End If
End Sub

Public Function StartsWith(s As String, prefix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        StartsWith = (Left(s, Len(prefix)) = prefix)
    Else
        StartsWith = (Left(LCase(s), Len(prefix)) = LCase(prefix))
    End If
End Function

' Determines whether a string ends with a given suffix.
Public Function EndsWith(s As String, suffix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        EndsWith = (Right(s, Len(suffix)) = suffix)
    Else
        EndsWith = (Right(LCase(s), Len(suffix)) = LCase(suffix))
    End If
End Function

' Splits a string on a given delimiter, trimming trailing and leading
' whitespace from each piece of the string.
Public Function SplitTrim(s As String, delim As String) As String()
    Dim arr() As String
    arr = Split(s, delim)
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        arr(i) = Trim(arr(i))
    Next
    
    SplitTrim = arr
End Function

' Trims a specified set of characters from the beginning and end
' of the given string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the beginning and end of the given string.
Public Function TrimChars(s As String, toTrim As String)
    TrimChars = TrimTrailingChars(TrimLeadingChars(s, toTrim), toTrim)
End Function

' Trims a specified set of characters from the beginning of the
' given string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the beginning of the given string.
Public Function TrimLeadingChars(s As String, toTrim As String)
    If s = "" Then
        TrimLeadingChars = ""
        Exit Function
    End If
    Dim i As Integer
    i = 1
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i <= Len(s)
        i = i + 1
    Wend
    TrimLeadingChars = Mid(s, i)
End Function

' Trims a specified set of characters from the end of the given
' string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the end of the given string.
Public Function TrimTrailingChars(s As String, toTrim As String)
    If s = "" Then
        TrimTrailingChars = ""
        Exit Function
    End If
    Dim i As Integer
    i = Len(s)
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i >= 1
        i = i - 1
    Wend
    TrimTrailingChars = Left(s, i)
End Function

Public Function ModuleExists(moduleName As String, Optional wb As Workbook) _
    As Boolean
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim c As Variant ' VBComponent
    
    On Error GoTo notFound
    Set c = wb.VBProject.VBComponents.Item(moduleName)
    ModuleExists = True
    Exit Function
    
notFound:
    ModuleExists = False
End Function

' Removes the VBA code module with the given name.
' @param wb: The workbook to remove the module from (defaults to the active
' workbook).
Public Sub RemoveModule(moduleName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then
        Err.Raise 32000, Description:= _
            "Module '" & moduleName & "' not found."
    End If
    Dim c As Variant ' VBComponent
    Set c = wb.VBProject.VBComponents.Item(moduleName)
    wb.VBProject.VBComponents.Remove c
    
    ' Sometimes the line above does not remove the module successfully.  When
    ' this happens, c.Name does not return an error - otherwise it does.
    On Error GoTo nameError
    Dim n As String
    n = c.Name
    On Error GoTo 0
    Err.Raise 32000, Description:= _
        "Failed to remove module '" & moduleName & "'.  Try again later."
    
nameError:
    ' Everything worked fine (the module was removed)
End Sub

' Exports a VBA code module to a text file.
' @param wb: The workbook that contains the module to export (defaults to the
' active workbook).
Public Sub ExportModule(moduleName As String, moduleFilename As String, _
    Optional wb As Workbook)
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then
        Err.Raise 32000, Description:= _
            "Module '" & moduleName & "' not found."
    End If
    wb.VBProject.VBComponents.Item(moduleName).Export moduleFilename
End Sub

' Imports a VBA code module from a text file.
' @param wb: The workbook that will receive the imported module (defaults to
' the active workbook).
Public Function ImportModule(moduleFilename As String, _
    Optional wb As Workbook) As VBComponent
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Set ImportModule = wb.VBProject.VBComponents.Import(moduleFilename)
End Function
