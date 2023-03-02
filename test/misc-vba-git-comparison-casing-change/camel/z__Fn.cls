VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "z__Fn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'zebra source MiscVBAFunctions#8e5e8f3

' Use a* so this is on top of the fn. MiscF library

Option Explicit

'*************** aFSO
Public fso As New fileSystemObject

'*************** aErrorEnums
Enum errNr
    '********************************************
    'These are internal error codes collected from
    'https://bettersolutions.com/vba/error-handling/error-codes.htm
    '********************************************
    returnWithoutGoSub = 3
    invalidProcedureCall = 5
    overflow = 6
    outOfMemory_ = 7
    subscriptOutOfRange = 9
    thisArrayIsFixedOrTemporarilyLocked = 10
    divisionByZero = 11
    typeMismatch = 13
    outOfStringSpace = 14
    expressionTooComplex = 16
    cantPerformRequestedOperation = 17
    userInterruptOccurred = 18
    resumeWithoutError = 20
    outOfStackSpace = 28
    subFunctionOrPropertyNotDefined = 35
    tooManyDLLApplicationClients = 47
    errorInLoadingDLL = 48
    badDLLCallingConvention = 49
    internalError = 51
    badFileNameOrNumber = 52
    fileNotFound = 53
    badFileMode = 54
    fileAlreadyOpen = 55
    deviceIOError = 57
    fileAlreadyExists = 58
    badRecordLength = 59
    diskFull = 61
    inputPastEndOfFile = 62
    badRecordNumber = 63
    tooManyFiles = 67
    deviceUnavailable = 68
    permissionDenied = 70
    diskNotReady = 71
    cantRenameWithDifferentDrive = 74
    pathFileAccessError = 75
    pathNotFound = 76
    objectVariableOrWithBlockVariableNotSet = 91
    forLoopNotInitialized = 92
    invalidPatternString = 93
    invalidUseOfNull = 94
    cantCallFriendProcedureOnAnObjectThatIsNotAnInstanceOfTheDefiningClass = 97
    systemDLLCouldNotBeLoaded = 298
    cantUseCharacterDeviceNamesInSpecifiedFileNames = 320
    invalidFileFormat = 321
    cantCreateNecessaryTemporaryFile = 322
    invalidFormatInResourceFile = 325
    dataValueNamedNotFound = 327
    illegalParameterCantWriteArrays = 328
    couldNotAccessSystemRegistry = 335
    activeXComponentNotCorrectlyRegistered = 336
    activeXComponentNotFound = 337
    activeXComponentDidNotRunCorrectly = 338
    objectAlreadyLoaded = 360
    cantLoadOrUnloadThisObject = 361
    activeXControlSpecifiedNotFound = 363
    objectWasUnloaded = 364
    unableToUnloadWithinThisContext = 365
    theSpecifiedFileIsOutOfDateThisProgramRequiresALaterVersion = 368
    theSpecifiedObjectCantBeUsedAsAnOwnerFormForShow = 371
    invalidPropertyValue = 380
    invalidPropertyarrayIndex = 381
    propertySetCantBeExecutedAtRunTime = 382
    propertySetCantBeUsedWithAReadonlyProperty = 383
    needPropertyArrayIndex = 385
    propertySetNotPermitted = 387
    propertyGetCantBeExecutedAtRunTime = 393
    propertyGetCantBeExecutedOnWriteonlyProperty = 394
    formAlreadyDisplayedCantShowModally = 400
    codeMustCloseTopmostModalFormFirst = 402
    permissionToUseObjectDenied = 419
    propertyNotFound = 422
    propertyOrMethodNotFound = 423
    objectRequired = 424
    invalidObjectUse = 425
    activeXComponentCantCreateObjectOrReturnReferenceToThisObject = 429
    classDoesntSupportAutomation = 430
    fileNameOrClassNameNotFoundDuringAutomationOperation = 432
    objectDoesntSupportThisPropertyOrMethod = 438
    automationError = 440
    connectionToTypeLibraryOrObjectLibraryForRemoteProcessHasBeenLost = 442
    automationObjectDoesntHaveADefaultValue = 443
    objectDoesntSupportThisAction = 445
    objectDoesntSupportNamedArguments = 446
    objectDoesntSupportCurrentLocaleSetting = 447
    namedArgumentNotFound = 448
    argumentNotOptionalOrInvalidPropertyAssignment = 449
    wrongNumberOfArgumentsOrInvalidPropertyAssignment = 450
    objectNotACollection = 451
    invalidOrdinal = 452
    specifiedDLLFunctionNotFound = 453
    codeResourceNotFound = 454
    codeResourceLockError = 455
    thisKeyIsAlreadyAssociatedWithAnElementOfThisCollection = 457
    variableUsesATypeNotSupportedInVisualBasic = 458
    thisComponentDoesntSupportEvents = 459
    invalidClipboardFormat = 460
    specifiedFormatDoesntMatchFormatOfData = 461
    cantCreateAutoRedrawImage = 480
    invalidPicture = 481
    printerError = 482
    printerDriverDoesNotSupportSpecifiedProperty = 483
    problemGettingPrinterInformationFromTheSystemMakeSureThePrinterIsSetUpCorrectly = 484
    invalidPictureType = 485
    cantPrintFormImageToThisTypeOfPrinter = 486
    cantEmptyClipboard = 520
    cantOpenClipboard = 521
    cantSaveFileToTEMPDirectory = 735
    searchTextNotFound = 744
    replacementsTooLong = 746
    outOfMemory = 31001
    noObject = 31004
    classIsNotSet = 31018
    unableToActivateObject = 31027
    unableToCreateEmbeddedObject = 31032
    errorSavingToFile = 31036
End Enum

'*************** Casing
' Uncomment and comment block to get casing back for the project
'
'
'Dim J
'Dim I
'Dim M
'Dim WB
'Dim WS
'Dim dict ' for now... :/
'Dim r ' for now... :/
'Dim Keys
'Dim Columns
'Dim Values
'Dim Column
'Dim ColumnName
'Dim Name
'Dim Path
'Dim After
'Dim Rows

'*************** MiscArray
Public Function errorToNullStringTransformation(tableArr() As Variant) As Variant
    ' Replaces all Errors in the input array with vbNullString.
    ' The input array is modified (pass by referance) and the function returns the array
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array that potentially contains error entries.
    '
    ' Returns:
    '   Array with the changed values.
    
    If miscArray_is2D(tableArr) Then
        errorToNullStringTransformation = miscArray_ErrorToNull2D(tableArr)
    Else
        errorToNullStringTransformation = miscArray_ErrorToNull1D(tableArr)
    End If
End Function

Public Function ensureDotSeparatorTransformation(tableArr() As Variant) As Variant
    ' Converts the decimal seperator in the float input to a "." for each entry in the input array
    ' and returns the result as an array of strings.
    ' Only works when converting from the system's decimal seperator.
    ' Custom seperators not supported.
    ' The input array is modified (pass by referance) and the function returns the array.
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array with float entries. Non numeric entries gets skipped.
    '
    ' Returns:
    '   Array with the changed string values.
    
    If miscArray_is2D(tableArr) Then
        ensureDotSeparatorTransformation = miscArray_EnsureDotSeparator2D(tableArr)
    Else
        ensureDotSeparatorTransformation = miscArray_EnsureDotSeparator1D(tableArr)
    End If
End Function

Public Function dateToStringTransformation(tableArr() As Variant, Optional fmt As String = "yyyy-mm-dd") As Variant
    ' Converts all Date/DateTime entries in the input array to string.
    ' The input array is modified (pass by referance) and the function returns the array.
    ' Functions for 1D and 2D arrays only.
    '
    ' Args:
    '   tableArr: Array with potential Date/DateTime entries.
    '   fmt: String format of the date that it must convert to. Default = "yyyy-mm-dd"
    '
    ' Returns:
    '   Array where the Date/DateTime entries have been converted.

    If miscArray_is2D(tableArr) Then
        dateToStringTransformation = miscArray_DateToString2D(tableArr, fmt)
    Else
        dateToStringTransformation = miscArray_DateToString1D(tableArr, fmt)
    End If
End Function

' Check if a collection is 1D or 2D.
' 3D is not supported
Private Function miscArray_is2D(arr As Variant)
    On Error GoTo err
    miscArray_is2D = (UBound(arr, 2) - LBound(arr, 2) > 1)
    Exit Function
err:
    miscArray_is2D = False
End Function

Private Function miscArray_dateToString(d As Date, fmt As String) As String
    miscArray_dateToString = Format(d, fmt)
End Function

Private Function miscArray_decStr(x As Variant) As String
     miscArray_decStr = CStr(x)

     'Frikin ridiculous loops for VBA
     If isNumeric(x) Then
        miscArray_decStr = replace(miscArray_decStr, Format(0, "."), ".")
        ' Format(0, ".") gives the system decimal separator
     End If

End Function

Private Function miscArray_ErrorToNull2D(tableArr As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(tableArr, 1) To UBound(tableArr, 1)
        For j = LBound(tableArr, 2) To UBound(tableArr, 2)
            If isError(tableArr(i, j)) Then ' set all error values to an empty string
                tableArr(i, j) = vbNullString
            End If
        Next j
    Next i
    miscArray_ErrorToNull2D = tableArr
End Function

Private Function miscArray_ErrorToNull1D(tableArr As Variant) As Variant
    Dim i As Long
    For i = LBound(tableArr) To UBound(tableArr)
        If isError(tableArr(i)) Then ' set all error values to an empty string
            tableArr(i) = vbNullString
        End If
    Next i
    miscArray_ErrorToNull1D = tableArr
End Function

Private Function miscArray_EnsureDotSeparator2D(tableArr As Variant) As Variant
    Dim i As Long, j As Long
    For i = LBound(tableArr, 1) To UBound(tableArr, 1)
        For j = LBound(tableArr, 2) To UBound(tableArr, 2)
            If isNumeric(tableArr(i, j)) Then ' force numeric values to use . as decimal separator
                tableArr(i, j) = miscArray_decStr(tableArr(i, j))
            End If
        Next j
    Next i
    miscArray_EnsureDotSeparator2D = tableArr
End Function

Private Function miscArray_EnsureDotSeparator1D(tableArr As Variant) As Variant
    Dim i As Long
    For i = LBound(tableArr) To UBound(tableArr)
        If isNumeric(tableArr(i)) Then ' force numeric values to use . as decimal separator
            tableArr(i) = miscArray_decStr(tableArr(i))
        End If
    Next i
    miscArray_EnsureDotSeparator1D = tableArr
End Function

Private Function miscArray_DateToString2D(tableArr As Variant, fmt As String) As Variant
    Dim i As Long, j As Long
    For i = LBound(tableArr, 1) To UBound(tableArr, 1)
        For j = LBound(tableArr, 2) To UBound(tableArr, 2)
            If isDate(tableArr(i, j)) Then ' format dates as strings to avoid some user's stupid default date settings
                tableArr(i, j) = miscArray_dateToString(CDate(tableArr(i, j)), fmt)
            End If
        Next j
    Next i
    miscArray_DateToString2D = tableArr
End Function

Private Function miscArray_DateToString1D(tableArr As Variant, fmt As String) As Variant
    Dim i As Long
    For i = LBound(tableArr, 1) To UBound(tableArr, 1)
        If isDate(tableArr(i)) Then ' format dates as strings to avoid some user's stupid default date settings
            tableArr(i) = miscArray_dateToString(CDate(tableArr(i)), fmt)
        End If
    Next i
    miscArray_DateToString1D = tableArr
End Function

'*************** MiscAssign
Public Function assign(ByRef var As Variant, ByRef val As Variant)
    ' Assign a value to a variable and also return that value. The goal of this function is to
    ' overcome the different `set` syntax for assigning an object vs. assigning a native type
    ' like Int, Double etc. Additionally this function has similar functionality to Python's
    ' walrus operator: https://towardsdatascience.com/the-walrus-operator-7971cd339d7d
    '
    ' Args:
    '   var: The input variable that could be an object.
    '   val: The value that the var input needs to be changed to.
    '
    ' Returns:
    '   The value from the input.
    
    If isObject(val) Then 'Object
        Set var = val
        Set assign = val
    Else 'Variant
        var = val
        assign = val
    End If
End Function

'*************** MiscCollection
Public Function min(ByVal col As collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If col Is Nothing Then
        err.raise number:=91, _
              description:="Collection input can't be empty"
    End If
    
    Dim entry As Variant
    min = col(1)
    
    For Each entry In col
        If entry < min Then
            min = entry
        End If
    Next entry
    
    
    
End Function

Public Function max(ByVal col As collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If col Is Nothing Then
        err.raise number:=91, _
              description:="Collection input can't be empty"
    End If
    
    max = col(1)
    Dim entry As Variant
    
    For Each entry In col
        If entry > max Then
            max = entry
        End If
    Next entry

End Function

Public Function mean(ByVal col As collection) As Variant
    ' Returns the mean value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The mean value of the collection.
    
    If col Is Nothing Then
        err.raise number:=91, _
              description:="Collection input can't be empty"
    End If

    mean = 0
    Dim entry As Variant
    
    For Each entry In col
        mean = mean + entry
    Next entry
    
    mean = mean / col.count
    
End Function

Public Function isValueInCollection(col As collection, val As Variant, Optional caseSensitive As Boolean = False) As Boolean
    ' Check if a value exists in the input Collection.
    '
    ' Args:
    '   col: Collection that potentially contains val
    '   val: The value to check for.
    '   CaseSensitive: Boolean entry to indicate if the comparison must be case sensitive.
    '
    ' Returns:
    '   True if val exists in the input Collection.
    
    Dim valI As Variant
    For Each valI In col
        ' only check if not an object:
        If Not isObject(valI) Then
            If caseSensitive Then
                isValueInCollection = valI = val
            Else
                isValueInCollection = VBA.lCase(valI) = VBA.lCase(val)
            End If
            ' exit if found
            If isValueInCollection Then Exit Function
        End If
    Next valI
End Function

'*************** MiscCollectionCreate
Public Function col(ParamArray args() As Variant) As collection
    ' Create a Collection from a list of entries.
    '
    ' Args:
    '   Args: list of entries that gets inserted into the Collection
    '
    ' Returns:
    '   Collection with the arguement values inserted.
    
    Set col = New collection
    Dim i As Long

    For i = LBound(args) To UBound(args)
        col.add args(i)
    Next

End Function

Public Function zip(ParamArray args() As Variant) As collection
    ' Standard zip function. Takes multiple Collections as an argument and
    ' group the matching index entries of each Collection into a new Collection.
    '
    ' Args:
    '   Args: Multiple Collections that gets grouped by index number.
    '
    ' Returns:
    '   A collection of collections containing the grouped entries.
    
    Dim i As Long
    Dim j As Long
    Dim m As Long
    
    m = -1
    For i = LBound(args) To UBound(args)
        If m = -1 Then
            m = args(i).count
        ElseIf args(i).count < m Then
            m = args(i).count
        End If
    Next i

    Set zip = New collection
    Dim iCol As collection
    For i = 1 To m
        Set iCol = New collection
        For j = LBound(args) To UBound(args)
            iCol.add args(j).item(i)
        Next j
        zip.add iCol
    Next i
End Function

'*************** MiscCollectionSort
Private Sub miscCollectionSort_TestBubbleSort()
    Dim coll As collection
    Set coll = col("variables10", "variables", "variables2", "variables_10", "variables_2")
    Set coll = bubbleSort(coll)
    
    Debug.Print coll(1), "variables"
    Debug.Print coll(2), "variables10" ' :/
    Debug.Print coll(3), "variables2" ' :/
    Debug.Print coll(4), "variables_10" ' :/
    Debug.Print coll(5), "variables_2" ' :/
    
End Sub

Public Function bubbleSort(coll As collection) As collection
    
    ' from: https://github.com/austinleedavis/VBA-utilities/blob/f23f1096d8df0dfdc740e5a3bec36525d61a3ffc/Collections.bas#L73
    ' this is an easy implementation but a slow sorting algorithm.
    ' do not use for large collections.
    '
    ' Args:
    '   coll: Unsorted Collection.
    '
    ' Returns:
    '   Sorted Collection
    
    Dim sortedColl As collection
    Set sortedColl = New collection
    Dim vItm As Variant
    ' copy the collection"
    For Each vItm In coll
        sortedColl.add vItm
    Next vItm

    Dim i As Long, j As Long
    Dim vTemp As Variant

    'Two loops to bubble sort
    For i = 1 To sortedColl.count - 1
        For j = i + 1 To sortedColl.count
            If sortedColl(i) > sortedColl(j) Then ' 1 = I is larger than J
                'store the lesser item
               assign vTemp, sortedColl(j) ' assign
                'remove the lesser item
               sortedColl.remove j
                're-add the lesser item before the greater Item
               sortedColl.add vTemp, , i
            End If
        Next j
    Next i
    
    Set bubbleSort = sortedColl
    
End Function

'*************** MiscCreateTextFile
Private Sub miscCreateTextFile_testCreateTextFile()
    createTextFile "foo", ThisWorkbook.path & "\tests\MiscCreateTextFile\test.txt"
    ' TODO: assertion
End Sub

Public Sub createTextFile(ByVal content As String, ByVal filePath As String)
    ' Creates a new / overwrites an existing text file with Content
    '
    ' Args:
    '   Content: Content that must be inserted into the file.
    '   FilePath: Path where the file will be created. The filename and extension must be included here.
    
    Dim oFile As Integer
    oFile = FreeFile
    
    Open filePath For Output As #oFile
        Print #oFile, content
    Close #oFile

End Sub

'*************** MiscDictionary
'@IgnoreModule ImplicitByRefModifier

Private Sub miscDictionary_testDictget()

    Dim d As dictionary
    Set d = dict("a", 2, "b", ThisWorkbook)
    
    
    Debug.Print dictget(d, "a"), 2 ' returns 2
    Debug.Print dictget(d, "b").Name, ThisWorkbook.Name ' returns the name of thisworkbook
    
    Debug.Print dictget(d, "c", vbNullString), vbNullString ' returns default value if key not found
    
    On Error Resume Next
        Debug.Print dictget(d, "c")
        Debug.Print err.number, 9 ' give error nr 9 if key not found
    On Error GoTo 0

End Sub

Public Function dictget(d As dictionary, key As Variant, Optional default As Variant = Empty) As Variant
    ' Return the entry in the input Dictionary at the given key. If the given key doesn't exist,
    ' the default value is returned if it's not empty. Else an error is raised.
    '
    ' Args:
    '   d: Dictionary to read the value from...
    '   key: The key value that gets used to return the input Dictionary's value with the matching key.
    '   default: The value that must be returned if the key doesn't exist in the Dictionary.
    '
    ' Returns:
    '   The Dictionary's entry or the default value.
    
    If d.exists(key) Then
        assign dictget, d.item(key)
        
    ElseIf Not isEmpty(default) Then
        assign dictget, default
        
    Else
        Dim errmsg As String
        On Error Resume Next
            errmsg = "Key "
            errmsg = errmsg & "`" & key & "` "
            errmsg = errmsg & "not in dictionary"
        On Error GoTo 0
        
        err.raise errNr.subscriptOutOfRange, , errorMessage(errNr.subscriptOutOfRange, errmsg)
    End If
End Function

'*************** MiscDictionaryCreate
Public Function dict(ParamArray args() As Variant) As dictionary
    ' Case sensitive dictionary
    '
    ' Args:
    '   Args: List of values that gets inserted into the Dictionary.
    '         All uneven entries are the keys and all even entries are the values for the matching keys.
    '
    ' Returns:
    '   The Dictionary
    
    Dim errmsg As String
    Set dict = New dictionary
    
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = LBound(args) To UBound(args)
        cnt = cnt + 1
        If (cnt Mod 2) = 0 Then GoTo cont

        If i + 1 > UBound(args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & args(i) & "`": On Error GoTo 0
            err.raise 9, , errmsg
        End If
        
        dict.add args(i), args(i + 1)
cont:
    Next i

End Function

Public Function dicti(ParamArray args() As Variant) As dictionary
    ' Case insensitive dictionary
    '
    ' Args:
    '   Args: List of values that gets inserted into the Dictionary.
    '         All uneven entries are the keys and all even entries are the values at its matching key.
    '
    ' Returns:
    '   The case insensitive Dictionary
    
    Dim errmsg As String
    Set dicti = New dictionary
    dicti.compareMode = textCompare
    
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    For i = LBound(args) To UBound(args)
        cnt = cnt + 1
        If (cnt Mod 2) = 0 Then GoTo cont

        If i + 1 > UBound(args) Then
            errmsg = "Dict construction is missing a pair"
            On Error Resume Next: errmsg = errmsg & " for key `" & args(i) & "`": On Error GoTo 0
            err.raise 9, , errmsg
        End If
        
        dicti.add args(i), args(i + 1)
cont:
    Next i

End Function

'*************** MiscDictsToTable
Public Function dictsToTable(tableDicts As collection, start_range As range, tableName As String) As listObject

    ' Converts a TableToDicts object back to a list object.
    '
    ' Args:
    '   TableDicts: Collection of dictionaries that represents a table [{"col1": 1, "col2": 2}, {"col1": 3, "col2": 4}]
    '   OutputStartRange: 1x1 range where the first header of the table should start
    '
    ' Returns:
    '   ListObject of the new table

    Dim numberOfRows, numberOfColumns, i, j As Integer
    Dim currentTable As listObject
    Dim dict As dictionary
    Dim columnNames() As Variant
    Dim columnNamesAsString As String
    Dim dictEntry As Variant
    
    numberOfRows = tableDicts.count
    numberOfColumns = tableDicts(1).count
    
    columnNames = tableDicts(1).keys()
    columnNamesAsString = join(columnNames, ",")

    For Each dict In tableDicts
        If dict.count <> numberOfColumns Then
            err.raise -997, , "Mismatch lengths for the dictionary entries. "
        End If
        
        For Each dictEntry In dict.keys()
        
            If (InStr(columnNamesAsString, dictEntry) = 0) Then
                err.raise -996, , "Mismatching dictionaries found. "
            End If
        Next dictEntry
    Next dict

    start_range.worksheet.listObjects.add(sourceType:=xlSrcRange, source:=start_range, _
    xlListObjectHasHeaders:=xlYes, tablestyleName:="TableStyleMedium2").Name = tableName

    Set currentTable = start_range.worksheet.listObjects(tableName)
    currentTable.listColumns.item(1).Name = tableDicts(1).keys()(0)
    
    ' Add column headers.
    For i = 1 To numberOfColumns - 1
        currentTable.listColumns.add.Name = tableDicts(1).keys()(i)
    Next i
    
    ' add values.
    For j = 1 To numberOfRows
        currentTable.listRows.add
        For i = 1 To numberOfColumns
            currentTable.listRows.item(j).range(i) = tableDicts(j)(columnNames(i - 1))
        Next i
    Next j
    
    Set dictsToTable = currentTable
End Function

'*************** MiscEnsureDictIUtil
Public Function ensureDictI(container As Variant) As Object
    ' Convert all Dicts in an object to case insensitive Dicts.
    ' The object can only contain Dicts and Collections.
    '
    ' Args:
    '   Container: Object that potentially contains Dicts.
    '
    ' Returns:
    '   A dict or Collection that potentially contains Dicts and/or Collections.
    
    Dim key As Variant
    Dim item As Variant
    
    If TypeOf container Is collection Then
        Dim c As collection
        Set c = New collection
        
        For Each item In container
            If TypeOf item Is collection Or TypeOf item Is dictionary Then
                c.add ensureDictI(item)
            Else
                c.add item
            End If
        Next item
        
        Set ensureDictI = c
        
    ElseIf TypeOf container Is dictionary Then
        Dim d As dictionary
        Set d = New dictionary
        d.compareMode = textCompare
        
        For Each key In container.keys
            If TypeOf container.item(key) Is collection Or TypeOf container.item(key) Is dictionary Then
                d.add key, ensureDictI(container.item(key))
            Else
                d.add key, container.item(key)
            End If
        Next key
        
        Set ensureDictI = d
    Else
    
        Dim errmsg As String
        errmsg = "ConvertToTextCompare only supports type 'Dictionary' and 'Collection'"
        On Error Resume Next: errmsg = errmsg & ". Got type '" & typeName(container) & "'": On Error GoTo 0
        err.raise 5, , errmsg
        
    End If
End Function

'*************** MiscErrorMessage
' #####################################################################
' ##### This module is version controlled in RunnerModule/Modules #####
' #####################################################################

'@Folder("error handling")

Public Function errorMessage(errorCode As Integer, _
                              Optional subMessage As String) As String
    ' Get the Error message for the given error code and return it. If the ErrorCode is not
    ' in the list of known error codes, "Unknown error" will be the Error message.
    ' SubMessage can be added that will be appended to the String that will be returned.
    '
    ' Args:
    '   ErrorCode: Error code to look for.
    '   SubMessage: Message to append to the returned error message.
    '
    ' Returns:
    '   String with the error message and the SubMessage.
    
    Dim m As String
    Select Case errorCode
        '**********************************************
        'VBA messages from https://bettersolutions.com/vba/error-handling/error-codes.htm
        '**********************************************
        Case 3: m = "Return without GoSub"
        Case 5: m = "Invalid procedure call"
        Case 6: m = "Overflow"
        Case 7: m = "Out of memory"
        Case 9: m = "Subscript out of range"
        Case 10: m = "This array is fixed or temporarily locked"
        Case 11: m = "Division by zero"
        Case 13: m = "Type mismatch"
        Case 14: m = "Out of string space"
        Case 16: m = "Expression too complex"
        Case 17: m = "Can't perform requested operation"
        Case 18: m = "User interrupt occurred"
        Case 20: m = "Resume without error"
        Case 28: m = "Out of stack space"
        Case 35: m = "Sub, Function, or Property not defined"
        Case 47: m = "Too many DLL application clients"
        Case 48: m = "Error in loading DLL"
        Case 49: m = "Bad DLL calling convention"
        Case 51: m = "Internal error"
        Case 52: m = "Bad file name or number"
        Case 53: m = "File not found"
        Case 54: m = "Bad file mode"
        Case 55: m = "File already open"
        Case 57: m = "Device I/O error"
        Case 58: m = "File already exists"
        Case 59: m = "Bad record length"
        Case 61: m = "Disk full"
        Case 62: m = "Input past end of file"
        Case 63: m = "Bad record number"
        Case 67: m = "Too many files"
        Case 68: m = "Device unavailable"
        Case 70: m = "Permission denied"
        Case 71: m = "Disk not ready"
        Case 74: m = "Can't rename with different drive"
        Case 75: m = "Path/File access error"
        Case 76: m = "Path not found"
        Case 91: m = "Object variable or With block variable not set"
        Case 92: m = "For loop not initialized"
        Case 93: m = "Invalid pattern string"
        Case 94: m = "Invalid use of Null"
        Case 97: m = "Can't call Friend procedure on an object that is not an instance of the defining class"
        Case 298: m = "System DLL could not be loaded"
        Case 320: m = "Can't use character device names in specified file names"
        Case 321: m = "Invalid file format"
        Case 322: m = "Can't create necessary temporary file"
        Case 325: m = "Invalid format in resource file"
        Case 327: m = "Data value named not found"
        Case 328: m = "Illegal parameter; can't write arrays"
        Case 335: m = "Could not access system registry"
        Case 336: m = "ActiveX component not correctly registered"
        Case 337: m = "ActiveX component not found"
        Case 338: m = "ActiveX component did not run correctly"
        Case 360: m = "Object already loaded"
        Case 361: m = "Can't load or unload this object"
        Case 363: m = "ActiveX control specified not found"
        Case 364: m = "Object was unloaded"
        Case 365: m = "Unable to unload within this context"
        Case 368: m = "The specified file is out of date. This program requires a later version"
        Case 371: m = "The specified object can't be used as an owner form for Show"
        Case 380: m = "Invalid property value"
        Case 381: m = "Invalid property-array index"
        Case 382: m = "Property Set can't be executed at run time"
        Case 383: m = "Property Set can't be used with a read-only property"
        Case 385: m = "Need property-array index"
        Case 387: m = "Property Set not permitted"
        Case 393: m = "Property Get can't be executed at run time"
        Case 394: m = "Property Get can't be executed on write-only property"
        Case 400: m = "Form already displayed; can't show modally"
        Case 402: m = "Code must close topmost modal form first"
        Case 419: m = "Permission to use object denied"
        Case 422: m = "Property not found"
        Case 423: m = "Property or method not found"
        Case 424: m = "Object required"
        Case 425: m = "Invalid object use"
        Case 429: m = "ActiveX component can't create object or return reference to this object"
        Case 430: m = "Class doesn't support Automation"
        Case 432: m = "File name or class name not found during Automation operation"
        Case 438: m = "Object doesn't support this property or method"
        Case 440: m = "Automation error"
        Case 442: m = "Connection to type library or object library for remote process has been lost"
        Case 443: m = "Automation object doesn't have a default value"
        Case 445: m = "Object doesn't support this action"
        Case 446: m = "Object doesn't support named arguments"
        Case 447: m = "Object doesn't support current locale setting"
        Case 448: m = "Named argument not found"
        Case 449: m = "Argument not optional or invalid property assignment"
        Case 450: m = "Wrong number of arguments or invalid property assignment"
        Case 451: m = "Object not a collection"
        Case 452: m = "Invalid ordinal"
        Case 453: m = "Specified DLL function not found"
        Case 454: m = "Code resource not found"
        Case 455: m = "Code resource lock error"
        Case 457: m = "This key is already associated with an element of this collection"
        Case 458: m = "Variable uses a type not supported in Visual Basic"
        Case 459: m = "This component doesn't support events"
        Case 460: m = "Invalid Clipboard format"
        Case 461: m = "Specified format doesn't match format of data"
        Case 480: m = "Can't create AutoRedraw image"
        Case 481: m = "Invalid picture"
        Case 482: m = "Printer error"
        Case 483: m = "Printer driver does not support specified property"
        Case 484: m = "Problem getting printer information from the system. Make sure the printer is set up correctly"
        Case 485: m = "Invalid picture type"
        Case 486: m = "Can't print form image to this type of printer"
        Case 520: m = "Can't empty Clipboard"
        Case 521: m = "Can't open Clipboard"
        Case 735: m = "Can't save file to TEMP directory"
        Case 744: m = "Search text not found"
        Case 746: m = "Replacements too long"
        Case 31001: m = "Out of memory"
        Case 31004: m = "No object"
        Case 31018: m = "Class is not set"
        Case 31027: m = "Unable to activate object"
        Case 31032: m = "Unable to create embedded object"
        Case 31036: m = "Error saving to file"
        Case 31037: m = "Error loading from file"
        Case Else: m = "Unknown error"
    End Select
    
    errorMessage = m & IIf(subMessage = "", "", ": " & subMessage)
    
    errorMessage = breakLines(errorMessage)
    
End Function

Public Function breakLines(subMessage As String) As String
    ' Split a word with 100+ characters into 3 lines and a word with 50+ words into 2 lines.
    '
    ' Args:
    '   SubMessage: String containing 1 or more words.
    '
    ' Returns:
    '   String with 50+ character words split between multiple lines.
    
    Dim msgArr() As String, i As Integer
    ' split by spaces
    msgArr = split(subMessage, " ")
    
    For i = LBound(msgArr) To UBound(msgArr)
        ' only split if one "word" is longer than 100 or 50 characters
        If Len(msgArr(i)) > 100 Then
            msgArr(i) = breakLines & Mid(msgArr(i), 1, 50) & chr(10) & Mid(msgArr(i), 51, 50) & chr(10) & right(msgArr(i), Len(msgArr(i)) - 50 * 2)
        ElseIf Len(msgArr(i)) > 50 Then
            msgArr(i) = Mid(msgArr(i), 1, 50) & chr(10) & right(msgArr(i), Len(msgArr(i)) - 50)
        End If
    Next i
    
    ' join back with spaces
    breakLines = join(msgArr, " ")
    
End Function

'*************** MiscExcel
Private Sub miscExcel_ModuleInitialize()
    Dim WB As Workbook
    Set WB = excelBook(fso.buildPath(ThisWorkbook.path, ".\tests\MiscExcel\MiscExcel23763464453.xlsx"), True)
    
End Sub

Public Function excelBook( _
      Optional path As String = "" _
    , Optional mustExist As Boolean = False _
    , Optional readOnly As Boolean = False _
    ) As Workbook
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/master/aa_py_xl/context.py
    ' Create an Excel Workbook with custom arguments.
    '
    ' Args:
    '   Path: Path to the file.
    '   MustExist: If True, the file must exist. If it doesn't an error is raised.
    '   ReadOnly: If True, the file is opened in readOnly mode.
    '
    ' Returns:
    '   The created/opened Workbook.
    
    If Len(path) = 0 Then
        If mustExist Then
            err.raise -997, , "Temp file can't have MustExist = True."
        End If
        If readOnly Then
            err.raise -996, , "Temp file can't open in ReadOnly mode."
        End If
        
        Set excelBook = workbooks.add
        Exit Function
    End If
    
    If fso.fileExists(path) Then
        Set excelBook = openWorkbook(path, readOnly)
        Exit Function
    End If
    
    If mustExist Then
        err.raise -999, , "FileNotFoundError: File '" & fso.getAbsolutePathName(path) & "' does not exist."
    End If
    
    If readOnly Then
        err.raise -998, , "File must exist to open in ReadOnly mode: File '" & fso.getAbsolutePathName(path) & "' does not exist."
    End If
    
    Set excelBook = workbooks.add
    excelBook.saveAs path
    
End Function

Public Function openWorkbook( _
      path As String _
    , Optional readOnly As Boolean = False _
    ) As Workbook
    ' Open a Workbook. An error is raised if a file with the same name is already open.
    ' If ReadOnly is True and the Workbook is already open but not in ReadOnly mode, an error is raised.
    '
    ' Args:
    '   Path: Path to the file that gets opened.
    '   ReadOnly: If True, the file gets opened in ReadOnly mode.
    '
    ' Returns:
    '   The opened Workbook.
    
    If hasKey(workbooks, fso.getFileName(path)) Then
        Set openWorkbook = workbooks(fso.getFileName(path))
        
        ' check if the workbook is actually the one specified in path
        ' use AbsolutePathName to remove any relative path references  (\..\ / \.\)
        If VBA.lCase(openWorkbook.fullName) <> VBA.lCase(fso.getAbsolutePathName(path)) Then
            Debug.Print fso.getAbsolutePathName(path)
            err.raise 457, , "Existing workbook with the same name is already open: '" & fso.getFileName(path) & "'"
        End If
        
        If readOnly And openWorkbook.readOnly = False Then
            err.raise -999, , "Workbook'" & fso.getFileName(path) & "' is already open and is not in ReadOnly mode. Only closed workbooks can be opened as readonly."
        End If
    Else
        Set openWorkbook = workbooks.Open(path, readOnly:=readOnly)
    End If
End Function

Public Function lastRow(WS As worksheet) As Integer
    ' Fetch the last row number that contains a value in any of the columns.
    ' Returns 0 if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last row number must be found
    '
    ' Returns:
    '   The last row number that contains a value in any of the columns.
    
    If worksheetFunction.countA(WS.cells) = 0 Then
        lastRow = 0
        Exit Function
    End If
    lastRow = WS.cells.find(what:="*", after:=WS.cells(1, 1), lookIn:=xlValues, lookat:=xlPart, searchOrder:=xlByRows, searchDirection:=xlPrevious).row
End Function

Public Function lastColumn(WS As worksheet) As Integer
    ' Fetch the last column number that contains a value in any of the rows.
    ' Returns 0 if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column number must be found
    '
    ' Returns:
    '   The last column number that contains a value in any of the rows.
    
    If worksheetFunction.countA(WS.cells) = 0 Then
        lastColumn = 0
        Exit Function
    End If
    lastColumn = WS.cells.find(what:="*", after:=WS.cells(1, 1), lookIn:=xlValues, lookat:=xlPart, searchOrder:=xlByColumns, searchDirection:=xlPrevious).column
End Function

Public Function lastCell(WS As worksheet) As range
    ' Get the last active cell using LastRow() and LastColumn() and returns it as a range.
    ' This function doesn't fetch the last cell with containing a value, but rather returns the
    ' the cell with the last active row and last active column.
    ' An unset Range object is returned if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column must be found.
    '
    ' Returns:
    '   The last cell in a worksheet as a range.
    
    Dim row, column As Integer
    row = lastRow(WS)
    column = lastColumn(WS)
    If row = 0 Or column = 0 Then
        Exit Function
    End If
    
    Set lastCell = WS.cells(row, column)
End Function

Public Function relevantRange(WS As worksheet) As range
    ' Get the relevant range of a Worksheet. This relevant range always
    ' starts at Cell("A1") and ends at the cell with the last active row and last active column.
    ' An unset Range object is returned if the Worksheet is empty.
    '
    ' Args:
    '   WS: The worksheet where the last column must be found.
    '
    ' Returns:
    '   The range of active cells in the selected Worksheet.
    
    Dim lastCellEntry As range
    Set lastCellEntry = lastCell(WS)
    
    If lastCellEntry Is Nothing Then
        Exit Function
    End If
    
    Set relevantRange = WS.range(WS.cells(1, 1), lastCellEntry)
End Function

'*************** MiscFreezePanes
Private Sub miscFreezePanes_test()
    Dim WS As worksheet
    Set WS = ThisWorkbook.worksheets(1)
    freezePanes WS.range("D6")
    
    
End Sub

Public Sub freezePanes(r As range)
    ' FreezePanes on the current active sheet. Removes FreezedPanes if it already exists.
    '
    ' Args:
    '   r: (row, column) cell where the FreezePanes should occur
    '
    
    Dim currentActiveSheet As worksheet
    Set currentActiveSheet = activeSheet
    
    Dim WS As worksheet
    Set WS = r.parent
    
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = application.screenUpdating
    application.screenUpdating = False
    With application.windows(WS.parent.Name)
        ' if existing freezed panes, remove them
        If .freezePanes = True Then
            .freezePanes = False
        End If
        application.GoTo WS.cells(1, 1) ' <- to ensure we don't hide the top/ left side of sheet
        ' Unfortunately, we have to do this :/
        application.GoTo r
        .freezePanes = True
    End With
    
    application.screenUpdating = currentScreenUpdating
    
    currentActiveSheet.activate
End Sub

Public Sub unFreezePanes(WS As worksheet)
    '
    '
    ' Args:
    '   WS: Worksheet where this function will execute.
    '

    Dim currentActiveSheet As worksheet
    Set currentActiveSheet = activeSheet
    
    ' Unfortunately, we have to do this :/
    WS.activate
    With application.windows(WS.parent.Name)
        .freezePanes = False
    End With
    
    currentActiveSheet.activate
End Sub

'*************** MiscGetUniqueItems
Private Sub miscGetUniqueItems_TestGetUniqueItems()
    Dim arr(3) As Variant
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "b"
    Debug.Print UBound(getUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(getUniqueItems(arr), 1), 3 ' zero index + case sensitive
    
    arr(0) = "a": arr(1) = "b": arr(2) = "c": arr(3) = "B"
    Debug.Print UBound(getUniqueItems(arr, False), 1), 2 ' zero index + case insensitive
    
    arr(0) = 1: arr(1) = 2: arr(2) = 3: arr(3) = 2
    Debug.Print UBound(getUniqueItems(arr), 1), 2 ' zero index
    
    arr(0) = 1: arr(1) = 1: arr(2) = "a": arr(3) = "a"
    Debug.Print UBound(getUniqueItems(arr), 1), 1 ' zero index
    
End Sub

Public Function getUniqueItems(arr() As Variant, _
            Optional caseSensitive As Boolean = True) As Variant
    'Return an array with unique values from the input array.
    '
    ' Args:
    '   arr: Array with potential duplicate entries.
    '   CaseSensitive: If true, the duplicate checks will be case sensitive.
    '
    ' Returns:
    '   An array with unique entries.
    
    If miscGetUniqueItems_ArrayLen(arr) = 0 Then
        getUniqueItems = Array()
    Else
        Dim d As New dictionary
        If Not caseSensitive Then
            d.compareMode = textCompare
        End If
        
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            If Not d.exists(arr(i)) Then
                d.add arr(i), arr(i)
            End If
        Next
        
        getUniqueItems = d.keys()
    End If
End Function

' Returns the number of elements in an array for a given dimension.
Private Function miscGetUniqueItems_ArrayLen(arr As Variant, _
    Optional dimNum As Integer = 1) As Long
    
    If isEmpty(arr) Then
        miscGetUniqueItems_ArrayLen = 0
    Else
        miscGetUniqueItems_ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function

'*************** MiscGroupOnIndentations
Public Sub groupRowsOnIndentations(r As range)
    ' groups the rows based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range or Rows that will be grouped.
    
    Dim ri As range
    For Each ri In r
        ri.entireRow.outlineLevel = ri.indentLevel + 1
    Next ri
    
End Sub

Public Sub groupColumnsOnIndentations(r As range)
    ' groups the columns based on indentations of the cells in the range
    '
    ' Args:
    '   r: Range of Columns that will be grouped.
    
    Dim ri As range
    For Each ri In r
        ri.entireColumn.outlineLevel = ri.indentLevel + 1
    Next ri
    
End Sub

Private Sub miscGroupOnIndentations_TestRemoveGroupings()
    
    ' Test rows
    removeRowGroupings ThisWorkbook.worksheets("GroupOnIndentations")
    ' Test columns
    removeColumnGroupings ThisWorkbook.worksheets("GroupOnIndentations")
End Sub

Public Sub removeRowGroupings(WS As worksheet)
    ' Remove Row Grouping from the selected Worksheet.
    '
    ' Args:
    '   WS: The workseheet where the grouping will be removed.
    
    Dim r As range
    Dim ri As range
    Set r = WS.usedRange ' todo: better way to find last "active" cell
    WS.outline.showLevels rowLevels:=8
    For Each ri In r.columns(1)
        ri.entireRow.outlineLevel = 1
    Next ri
End Sub

Public Sub removeColumnGroupings(WS As worksheet)
    ' Remove Column Grouping from the selected Worksheet.
    '
    ' Args:
    '   WS: The workseheet where the grouping will be removed.
    
    Dim r As range
    Dim ri As range
    Set r = WS.usedRange ' todo: better way to find last "active" cell
    WS.outline.showLevels columnlevels:=8
    For Each ri In r.rows(1)
        ri.entireColumn.outlineLevel = 1
    Next ri
End Sub

'*************** MiscHasKey
'@IgnoreModule ImplicitByRefModifier

Private Sub miscHasKey_TestHasKey()

    Dim c As New collection
    c.add "a", "a"
    c.add col("x", "y", "z"), "b"
    
    Debug.Print vbLf & "*********** TestHasKey tests ***********"
    Debug.Print True, hasKey(c, "a") ' True for scalar
    Debug.Print True, hasKey(c, "b") ' True for object
    Debug.Print True, hasKey(c, "A") ' True (even though case insensitive???)

    Debug.Print True, hasKey(workbooks, ThisWorkbook.Name) ' True for non-collection type collections
    
    Dim d As New dictionary
    d.add "a", "a"
    d.add "b", col("x", "y", "z")
    
    Debug.Print True, hasKey(d, "a") ' True for scalar
    Debug.Print True, hasKey(d, "b") ' True for object
    Debug.Print False, hasKey(d, "A") ' False - case sensitive by default
    
    Dim dObj As Object
    Set dObj = createObject("Scripting.Dictionary")
    
    dObj.add "a", "a"
    dObj.add "b", col("x", "y", "z")
    
    Debug.Print True, hasKey(dObj, "a") ' True for scalar
    Debug.Print True, hasKey(dObj, "b") ' True for object
    Debug.Print False, hasKey(dObj, "A") ' False - case sensitive by default
    
    ' Errors
    On Error Resume Next
        err.number = 0
        hasKey ThisWorkbook, "A"
        Debug.Print 9, err.number ' WorkBook doesn't have keys/items
        
        err.number = 0
        hasKey 5, "A"
        Debug.Print 9, err.number ' Variant doesn't have keys/items
    On Error GoTo 0

End Sub

Public Function hasKey(container As Variant, key As Variant) As Boolean
    ' Checks whether a key exists in an existing container
    ' The container can be a `Collection`, `Dictionary` or any
    ' built-in Dictionary-like object. For example `ThisWorkbook.Sheets`
    '
    ' Args:
    '     Container: The container in which to look for the key
    '     key: The key to look for in the container
    '
    ' Returns:
    '     True for success, False otherwise
    
    
    Dim errX As Integer
    Dim hasKeyFlag As Boolean
    Dim emptyFlag As Boolean
    
    ' First try .HasKey method on the object
    On Error Resume Next
        err.number = 0
        hasKeyFlag = container.exists(key)
        errX = err.number
    On Error GoTo 0
    If errX = 0 Then
        hasKey = hasKeyFlag
        Exit Function
    End If
    
    
    ' Then test with .Item method
    emptyFlag = False
    On Error Resume Next
        err.number = 0
        emptyFlag = typeName(container.item(key)) = "Empty"
        errX = err.number
    On Error GoTo 0
    
    If errX = 0 Then ' No error trying to Access Key via .Item
        hasKey = Not emptyFlag
        Exit Function
    ElseIf errX <> 424 And errX <> 438 Then ' Retrieval Error, but .Item is correct access method stil. 424: Method not exist; 438: Compilation error
        hasKey = False
        Exit Function
    End If
    
    
    ' Then test with bracketed access, like ()
    emptyFlag = False
    On Error Resume Next
        err.number = 0
        emptyFlag = typeName(container(key)) = "Empty"
        errX = err.number
    On Error GoTo 0
    
    If errX = 0 Then ' No error trying to Access Key via ()
        hasKey = Not emptyFlag
        Exit Function
    ElseIf errX <> 424 And errX <> 438 And errX <> 13 Then ' Retrieval Error, but () is correct access method stil. 424: Method not exist; 438: Compilation error; 13: Variant bracketed ()
        hasKey = False
        Exit Function
    End If

    
    Dim errmsg As String
    On Error Resume Next
        errmsg = "Object"
        errmsg = errmsg & " of type '" & typeName(container) & "'"
        errmsg = errmsg & " have neither '.Exists' method, nor bracketed indexing '()', nor '.Item' method"
    On Error GoTo 0
    err.raise 9, , errmsg
    
End Function

'*************** MiscNewKeys
Public Function newSheetName(Name As String, Optional WB As Workbook)
    ' this module is used to generate new keys to a container (collections, dict, sheets, etc)
    ' Use case is when we want to create a new sheet, but
    ' want to ensure we don't give a name that already exists in the workbook
    '
    ' Args:
    '   Name: Name of the Sheet.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   The unique name of the container.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' max 31 characters
    newSheetName = Left(Name, 31)

    If Not hasKey(WB.sheets, newSheetName) Then
        ' sheet name doesn't exist, so we can continue
        Exit Function
    Else
        newSheetName = getNewKey(Name, WB.sheets, 31)
    End If
End Function

Private Sub miscNewKeys_TestGetNewKey()

    Dim c As New collection
    Dim i As Long
    
    c.add "bla", "name"
    For i = 1 To 100
        c.add "bla", "name" & i
    Next i
    
    Debug.Print getNewKey("name", c), "name101"
    Debug.Print getNewKey("NewName", c), "NewName"

End Sub

Public Function getNewKey(Name As String, container As Variant, Optional maxLength As Long = -1, Optional depth As Long = 0) As String
    ' get a key that does not exists in a container (dict or collection)
    ' we keep appending, 1, 2, 3, ..., 10, 11 until the key is unique
    ' MaxLength is used when the key has a restriction on the maximum length
    ' for example sheet names can only be 31 characters long
    '
    ' Args:
    '   Name: Name of the key
    '   Container: Container containing the existing keys
    '   MaxLength: Maximum length of the resulting key.
    '   depth: Starting number to append to the key, while searching for a unique key.
    '
    ' Returns:
    '   The unique key
    
    If maxLength = -1 Then
        getNewKey = Name
    Else
        getNewKey = Left(Name, maxLength)
    End If
    
    If Not hasKey(container, getNewKey) Then
        ' Key is "New" and we don't need further iteration
        Exit Function
    Else
        ' 31 max characters for sheet name
        depth = depth + 1
        If maxLength = -1 Then
            getNewKey = getNewKey & depth
        Else
            getNewKey = Left(getNewKey, maxLength - Len(CStr(depth))) & depth
        End If
        
        If Not hasKey(container, getNewKey) Then
            Exit Function
        End If
        
        getNewKey = getNewKey(Name, container, maxLength, depth)
    End If
End Function

'*************** MiscOs
Public Function path(ParamArray paths() As Variant) As String
    ' Combines folder pathes and the name of folders or a file and
    ' returns the combination with valid path separators.
    '
    ' Args:
    '   entries: The folder pathes and the name of folders or a file to be combined.
    '
    ' Returns:
    ' The combination of paths with valid path separators.
    
    Dim entry As Variant
    
    path = paths(0)

    For entry = LBound(paths) + 1 To UBound(paths)
        path = fso.buildPath(path, paths(entry))
    Next
    
End Function

'*************** MiscPowerQuery
' Helpful functions to help with Power Query manipulations in VBA

Private Sub miscPowerQuery_MiscPowerQueryTests()
    Debug.Print doesQueryExist("foo"), False
End Sub

Public Function doesQueryExist(ByVal queryName As String, Optional WB As Workbook) As Boolean
    ' Check if a Query exists in the given Workbook.
    '
    ' Args:
    '   queryName: Name of the Query to look for.
    '   WB: Name of the WorkBook to look in.
    '
    ' Returns:
    '   True if the Query exists, False otherwise.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Helper function to check if a query with the given name already exists
    Dim qry As workbookQuery
    For Each qry In WB.queries
        If (qry.Name = queryName) Then
            doesQueryExist = True
            Exit Function
        End If
    Next
    doesQueryExist = False
End Function

Public Function getQuery(Name As String, Optional WB As Workbook) As workbookQuery
    ' Return the desired Query if it exists. If the Query doesn't exist, an error is raised.
    '
    ' Args:
    '   Name: Name of the Query to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The desired Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim qry As workbookQuery
    For Each qry In WB.queries
        If qry.Name = Name Then
            Set getQuery = qry
            Exit Function
        End If
    Next qry
    
    err.raise 999, , "Query " & Name & " does not exist"
    
End Function

Public Function getQueryFormula(Name As String, Optional WB As Workbook) As String
    ' Returns the Power Query M formula of a WorkbookQuery
    '
    ' Args:
    '   Name: Name of the Query to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The Power Query M formula of the WorkbookQuery
    getQueryFormula = getQuery(Name, WB).formula
End Function

Public Function updateQuery(Name As String, queryFormula As String, Optional WB As Workbook) As workbookQuery
    ' Update the selected Query. If the Query doesn't exist, a new Query is added.
    '
    ' Args:
    '   Name: Name of the Query.
    '   queryFormula: New Formula of the Query.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   Updated or new Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a query to the new formula
    ' if the query doesn't exist, a new one is created
    
    If doesQueryExist(Name, WB) Then
        Set updateQuery = getQuery(Name, WB)
        updateQuery.formula = queryFormula
    Else
        Set updateQuery = WB.queries.add(Name, queryFormula)
    End If
    
End Function

Public Function updateQueryAndRefreshListObject(Name As String, queryFormula As String, Optional WB As Workbook) As workbookQuery
    ' Update the selected Query and refresh the list of objects.
    '
    ' Args:
    '   Name: Name of the Query to update.
    '   queryFormula: New Formula of the Query.
    '   WB: The selected Workbook.
    '
    ' Returns:
    '   Updated or new Query.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' updates a power query query
    ' Also waits for the query to refresh before continuing the code
    
    ' assumes the ListObject and Query has the same name
    Set updateQueryAndRefreshListObject = updateQuery(Name, queryFormula, WB)
    
    waitForListObjectRefresh Name, WB
    
End Function

Public Sub waitForListObjectRefresh(Name As String, Optional WB As Workbook)
    ' Refresh elements in the QueryTable.
    '
    ' Args:
    '   Name: Name of the ListObject.
    '   WB: Name of the WorkBook.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    ' Refreshes the query before continuing the code
    
    Dim LO As listObject
    Set LO = getLO(Name, WB)
    Dim bGRefresh As Boolean
    With LO.queryTable
        bGRefresh = .backgroundQuery
        .backgroundQuery = False
        .refresh
        .backgroundQuery = bGRefresh
    End With
    
End Sub

Public Sub loadToWorkbook(queryName As String, Optional WB As Workbook)
    ' loads a query to a sheet in the workbook
    '
    ' Args:
    '   queryName: Name of the query to load to the WorkBook
    '   WB: Name of the WorkBook.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim LO As listObject
    If hasLO(queryName, WB) Then
        Set LO = getLO(queryName, WB)
        LO.refresh
    Else
        Dim WS As worksheet
        Set WS = WB.worksheets.add(after:=activeSheet)
        WS.Name = newSheetName(queryName, ThisWorkbook)
        
        With WS.listObjects.add(sourceType:=0, source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""" _
            , destination:=range("$A$1")).queryTable
            .commandType = xlCmdSql
            .commandText = Array("SELECT * FROM [" & queryName & "]")
            .rowNumbers = False
            .fillAdjacentFormulas = False
            .preserveFormatting = True
            .refreshOnFileOpen = False
            .backgroundQuery = True
            .refreshStyle = xlInsertDeleteCells
            .savePassword = False
            .saveData = True
            .adjustColumnWidth = True
            .refreshPeriod = 0
            .preserveColumnInfo = True
            .listObject.displayName = queryName
            .refresh backgroundQuery:=False
        End With
        
    End If
    
End Sub

Public Function addToWorkbookConnections(query As workbookQuery, Optional WB As Workbook) As workbookConnection
    ' adds a query to workbookconnections so that it can be used in pivot tables
    '
    ' Args:
    '   Query: Query that gets added to the workbookconnections.
    '   WB: Name of the WorkBook
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Dim connectionName As String, commandString As String, commandText As String, commandType
    connectionName = "Query - " & query.Name
    commandString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name & ";Extended Properties="""""
    commandText = "SELECT * FROM [" & query.Name & "]"
    commandType = 2
    
    ' This code loads the query to the workbook connections
    If hasKey(WB.connections, connectionName) Then
        Set addToWorkbookConnections = WB.connections(connectionName)
        addToWorkbookConnections.oLEDBConnection.connection = commandString
        addToWorkbookConnections.oLEDBConnection.commandText = commandText
        addToWorkbookConnections.oLEDBConnection.commandType = commandType
    Else
        Set addToWorkbookConnections = _
        WB.connections.add2(connectionName, _
            "Connection to the '" & query.Name & "' query in the workbook.", _
            commandString _
            , commandText, commandType)
        ' should not be loaded to the data model, else we cannot link two pivots to the same cache linking from this query
    End If

End Function

Public Sub refreshAllQueriesAndPivots(Optional WB As Workbook)
    ' Refresh all Queries and Pivots.
    '
    ' Args:
    '   WB: Name of the WorkBook
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    WB.refreshAll
End Sub

'*************** MiscRangeToArray
Public Function rangeToArray(r As range, _
                Optional ignoreEmptyInFlatArray As Boolean) As Variant()
    ' Converts a range to a normalized array.
    ' vectors allocated to 1-dimensional arrays
    ' tables allocated to 2-dimensional array
    '
    ' Args:
    '   r: Range to be converted to an array.
    '   IgnoreEmptyInFlatArray: If True, skip over empty results.
    '
    ' Returns:
    '   The normalized array.
    
    If r.cells.count = 1 Then
        rangeToArray = Array(r.value)
    ElseIf r.rows.count = 1 Or r.columns.count = 1 Then
        rangeToArray = rangeTo1DArray(r, ignoreEmptyInFlatArray)
    Else
        rangeToArray = rangeTo2DArray(r)
    End If
End Function

Public Function rangeTo1DArray( _
              r As range _
            , Optional ignoreEmpty As Boolean = True _
            ) As Variant()
    ' currently does the same as rangeToArray, just named better and is more efficient
    ' instead of reading from memory for every range item, we read it in only once
    '
    ' Args:
    '   r: Range to be converted to an array.
    '   IgnoreEmpty: If True, skip over empty results.
    '
    ' Returns:
    '   The normalized array.
    
    Dim arr() As Variant ' the output array
    ReDim arr(r.cells.count - 1)
    
    Dim values() As Variant ' values of the whole range
    If r.cells.count = 1 Then
        arr(0) = r.value
        rangeTo1DArray = arr
        Exit Function
    End If
    
    values = r.value
    Dim i As Long
    Dim j As Long
    Dim counter As Long
    counter = 0
    For i = LBound(values, 1) To UBound(values, 1) ' rows
        For j = LBound(values, 2) To UBound(values, 2) ' columns
            If isError(values(i, j)) Then
                ' if error, we cannot check if empty, we need to add it
                arr(counter) = values(i, j)
                counter = counter + 1
            ElseIf values(i, j) = vbNullString And ignoreEmpty Then
                ReDim Preserve arr(UBound(arr) - 1) ' when there is an empty cell, just reduce array size by 1
            Else
                arr(counter) = values(i, j)
                counter = counter + 1
            End If
        Next j
    Next i
    
    rangeTo1DArray = arr
    
End Function

Public Function rangeTo2DArray(r As range) As Variant()
    ' ensure a range is converted to a 2-dimensional array
    ' special treatment on edge cases where a range is a 1x1 scalar
    '
    ' Args:
    '   r: Range to be converted to an array.
    '
    ' Returns:
    '   2D array.
    
    If r.cells.count = 1 Then
        Dim arr_single() As Variant
        ReDim arr_single(1 To 1, 1 To 1) ' make it base 1, similar to what .value does for non-scalars
        arr_single(1, 1) = r.value
        rangeTo2DArray = arr_single
        Exit Function
    End If
    
    Dim values() As Variant ' values of the whole range
    values = r.value

    Dim arr() As Variant ' the output array
    ReDim arr(UBound(values, 1) - LBound(values, 1), UBound(values, 2) - LBound(values, 2))
    Dim i As Long
    Dim j As Long
    Dim i_start As Long
    Dim j_start As Long
    i_start = LBound(values, 1)
    j_start = LBound(values, 2)
    For i = LBound(values, 1) To UBound(values, 1) ' rows
        For j = LBound(values, 2) To UBound(values, 2) ' columns
            arr(i - i_start, j - j_start) = values(i, j)
        Next j
    Next i
    rangeTo2DArray = arr
    
End Function

'*************** MiscRemoveGridLines
Public Sub removeGridLines(WS As worksheet)
    ' Remove all GridLines from the selected Worksheet.
    '
    ' Args:
    '   WS: Selected WorkSheet.
    
    Dim view As worksheetView
    For Each view In WS.parent.windows(1).sheetViews
        If view.sheet.Name = WS.Name Then
            view.displayGridlines = False
            Exit Sub
        End If
    Next
End Sub

'*************** MiscString
Public Function randomString(length As Variant)
    ' Create a random string containing hex characters only.
    ' (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, C, D, E, F)
    '
    ' Args:
    '   length: Number of characters that the string must have.
    '
    ' Returns:
    '   The Random string.
    
    Dim s As String
    While Len(s) < length
        s = s & hex(rnd * 16777216)
    Wend
    randomString = Mid(s, 1, length)
End Function

'*************** MiscTables
Public Function hasLO(Name As String, Optional WB As Workbook) As Boolean
    ' Check if the selected WorkBook contains a ListObject with the input name.
    '
    ' Args:
    '   Name: Name of the ListObject to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   True if the ListObject exists.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As worksheet
    Dim LO As listObject
    
    For Each WS In WB.worksheets
        For Each LO In WS.listObjects
            If VBA.lCase(Name) = VBA.lCase(LO.Name) Then
                hasLO = True
                Exit Function
            End If
        Next LO
    Next WS
    
    hasLO = False

End Function

Public Function getLO(Name As String, Optional WB As Workbook) As listObject
    ' get list object only using it's name from within a workbook
    '
    ' Args:
    '   Name: Name of the ListObject to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The ListObject if it exists. An error is raised if it doesn't exist.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    Dim WS As worksheet
    Dim LO As listObject
    
    For Each WS In WB.worksheets
        For Each LO In WS.listObjects
            If VBA.lCase(Name) = VBA.lCase(LO.Name) Then
                Set getLO = LO
                Exit Function
            End If
        Next LO
    Next WS
    
    If getLO Is Nothing Then
        ' 9: Subscript out of range
        err.raise errNr.subscriptOutOfRange, , errorMessage(errNr.subscriptOutOfRange, "List object '" & Name & "' not found in workbook '" & WB.Name & "'")
    End If

End Function

Private Sub miscTables_TestTableToArray()
    tableToArray "foo"
End Sub

Public Function tableToArray( _
      Name As String _
    , Optional WB As Workbook _
    ) As Variant()
    ' Return an Array of the input table.
    '
    ' Args:
    '   Name: Name of the table to look for.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   2D array of the selected Table.
    
    tableToArray = rangeTo2DArray(tableRange(Name, WB))
    
End Function

Public Function tableRange( _
        Name As String _
      , Optional WB As Workbook _
      ) As range
    
    'Returns the range (including headers of a table named `Name` in workbook `WB`): _
    - It first looks for a list object called `Name` _
      - If the `.DataBodyRange` property is nothing the table range will only be the headers _
    - Then it looks for a named range in the Workbook scope called `Name` and returns the _
      range this named range is referring to _
    - Then it looks for a worksheet scoped named range called `Name`. The first occurrence _
      will be returned _
    If no tables found, a `SubscriptOutOfRange` error (9) is raised _
    The name of the table to be found is case insensitive
    '
    ' Args:
    '   Name: Name of the table to look for.
    '   WB: Selected Workbook.
    '
    ' Returns:
    '   Range of the cells in the selected Table.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    If hasLO(Name, WB) Then
        Dim LO As listObject
        Set LO = getLO(Name, WB)
        If LO.dataBodyRange Is Nothing Then
            Set tableRange = LO.headerRowRange
        Else
            Set tableRange = LO.range
        End If
        Exit Function
    End If
    
    If hasKey(WB.names, Name) Then
        Set tableRange = WB.names(Name).refersToRange
        Exit Function
    End If
    
    Dim WS As worksheet
    ' this will find the first occurrence of the table called 'Name'
    For Each WS In WB.worksheets
        If hasKey(WS.names, Name) Then
            Set tableRange = WS.names(Name).refersToRange
            Exit Function
        End If
    Next WS
    
    err.raise errNr.subscriptOutOfRange, , errorMessage(errNr.subscriptOutOfRange, "Table '" & Name & "' not found in workbook '" & WB.Name & "'")
    
End Function

Public Function getAllTables(WB As Workbook) As collection
    Set getAllTables = New collection
    ' Returns all tables in a workbook
    '
    ' Args:
    '   WB: The selected WorkBook
    '
    ' Returns:
    '   All tables in the selected WorkBook.
    
    Dim WS As worksheet
    Dim LO As listObject
    For Each WS In WB.worksheets
        For Each LO In WS.listObjects
            getAllTables.add LO.Name
        Next LO
    Next WS
    
    Dim Name As Name
    For Each Name In WB.names
        getAllTables.add Name.Name
    Next Name
    
    For Each WS In WB.worksheets
        For Each Name In WS.names
            ' remove the sheetname prefix to get the table name
            getAllTables.add Mid(Name.Name, InStr(Name.Name, "!") + 1)
        Next Name
    Next WS
    
End Function

Function tableColumnToArray(tableDicts As collection, columnName As String) As Variant()
    ' Append the selected key's value from each Dict in the input Collection to a 1-dimensional array
    '
    ' Args:
    '   TableDicts: A collection of Dicts.
    '   ColumnName: Name of the column that will be returned as a 1-D array.
    '
    ' Returns:
    '   1-D array of the selected column.
    
    Dim arr() As Variant
    ReDim arr(tableDicts.count - 1) ' zero indexed
    Dim dict As dictionary
    Dim counter As Long
    For Each dict In tableDicts
        arr(counter) = dictget(dict, columnName)
        counter = counter + 1 ' zero indexing
    Next dict
    
    tableColumnToArray = arr
End Function

'*************** MiscTableToDicts
Private Sub miscTableToDicts_TableToDictsTest()
    Dim dicts As collection
    Set dicts = tableToDicts("TableToDictsTestData")
    ' read row 2 in column "b":
    Debug.Print dicts(2)("b"), 5
End Sub

Public Function tableToDictsLogSource( _
          tableName As String _
        , Optional WB As Workbook _
        , Optional columns As collection _
        ) As collection
    
    'Similar to TableToDicts, but also stores the source of each row _
    in a dictionary with key `__source__`
    'The `__source__` object contains the following keys: _
     - `Workbook`: the Workbook object with the table _
     - `Table`: the name of the table within the workbook _
     - `RowIndex`: the row index of the current entry of the table
    '
    ' Args:
    '   TableName: Name of the table to convert to Dicts.
    '   WB: Selected WorkBook
    '   Columns: Columns to include in the Dicts.
    '
    ' Returns:
    '   The collection of Dicts containing the info as well as the source of each row.
    
    Set tableToDictsLogSource = tableToDicts(tableName, WB, columns)
    Dim dict As dictionary
    Dim rowIndex As Long
    rowIndex = 0
    For Each dict In tableToDictsLogSource
        rowIndex = rowIndex + 1
        dict.add "__source__", dicti("Workbook", WB, "Table", tableName, "RowIndex", rowIndex)
    Next dict
End Function

Public Function tableToDicts( _
          tableName As String _
        , Optional WB As Workbook _
        , Optional columns As collection _
        ) As collection
    
    ' Inspiration: https://github.com/AutoActuary/aa-py-xl/blob/8e1b9709a380d71eaf0d59bd0c2882c8501e9540/aa_py_xl/data_util.py#L21
    ' Convert a Table to a Collection of Dicts.
    '
    ' Args:
    '   TableName: Name of the Selected Table.
    '   WB: Selected WorkBook
    '   Columns: Columns to be added to the Dicts.
    '
    ' Returns:
    '   A collection of Dictionaries.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    Set tableToDicts = New collection
    
    Dim d As dictionary
    
    Dim i As Long
    Dim j As Long
    Dim tableData() As Variant
    tableData = tableToArray(tableName, WB)
    
    For i = LBound(tableData, 1) + 1 To UBound(tableData, 1)
        Set d = New dictionary
        d.compareMode = textCompare ' must be case insensitive
        
        If columns Is Nothing Then
            For j = LBound(tableData, 2) To UBound(tableData, 2)
                d.add tableData(0, j), tableData(i, j)
            Next j
        Else
            Dim columnName As Variant
            Dim column As Variant
            
            For j = LBound(tableData, 2) To UBound(tableData, 2)
                columnName = tableData(LBound(tableData, 2), j)
                If isValueInCollection(columns, columnName) Then
                    d.add columnName, tableData(i, j)
                End If
            Next j
        End If
        
        tableToDicts.add d
    Next i
    
End Function

Private Function miscTableToDicts_TestGetTableRowIndex()
    Dim table As collection
    Set table = col(dicti("a", 1, "b", 2), dicti("a", 3, "b", 4), dicti("a", "foo", "b", "bar"))
    Debug.Print getTableRowIndex(table, col("a", "b"), col(3, 4)), 2
    Debug.Print getTableRowIndex(table, col("a", "b"), col("foo", "bar")), 3
End Function

Public Function tableLookupValue( _
        table As Variant _
      , columns As collection _
      , values As collection _
      , valueColName As String _
      , Optional default As Variant = Empty _
      , Optional WB As Workbook _
      ) As Variant
    ' Returns the value from the ValueColName column in a TableToDicts object _
      given the value In the lookup column _
      A default value can be assigned For when no lookup Is found _
      Otherwise it returns a runtime Error
    '
    ' Args:
    '   Table: Selected table.
    '   Columns: Collection of selected Column names.
    '   Values: Values from the lookup column
    '   ValueColName: Column name that gets used to fetch values from.
    '   default: Value to be used when no value has been found.
    '   WB: Selected workbook.
    '
    ' Returns:
    '   Value from the ValueColName column.
    
    If WB Is Nothing Then Set WB = ThisWorkbook
    
    ' for when GetTableRowIndex fails
    If Not isEmpty(default) Then On Error GoTo setDefault
    
    Dim dict As dictionary
    Set dict = miscTableToDicts_EnsureTableDicts(table, WB)(getTableRowIndex(table, columns, values, WB))
    tableLookupValue = dictget(dict, valueColName, default)
    
    Exit Function
setDefault:
    tableLookupValue = default
    
End Function

Public Function getTableRowRange( _
      tableName As String _
    , columns As collection _
    , values As collection _
    , Optional WB As Workbook _
    ) As range
    ' Given a table name, Columns and Values to match _
      this function returns the row in which these values matches
    ' Comparison is case sensitive
    ' If no match is found, a runtime error is raised
    '
    ' Args:
    '   TableName: Name of the Table
    '   Columns: Collection of Column names.
    '   Values: Values to match agains.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The row in which the vales matches the comparison.
    
    Dim rowNumber As Long
    rowNumber = getTableRowIndex(tableName, columns, values, WB) ' this will throw a runtime error if not found
    
    Dim tableR As range
    Set tableR = tableRange(tableName, WB)
    
    ' Intersect of table range and entirerow
    ' +1 as header is not included in GetTableRowIndex
    Set getTableRowRange = intersect(tableR, tableR(rowNumber + 1, 1).entireRow)
    
End Function

Public Function getTableColumnRange( _
      tableName As String _
    , column As String _
    , Optional WB As Workbook _
    ) As range
    ' Returns the range of a table's column, including the header
    '
    ' Args:
    '   TableName: Name of the Table.
    '   Columns: Name of the column.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   Range of cells for the selected table's column.
    
    Dim tableR As range
    Set tableR = tableRange(tableName, WB)
    
    Dim i As Long
    For i = 1 To tableR.columns.count
        If VBA.lCase(tableR(1, i).value) = VBA.lCase(column) Then
            GoTo found
        End If
    Next i
    
    err.raise errNr.subscriptOutOfRange, , errorMessage(errNr.subscriptOutOfRange, "Column '" & column & "' not found in table '" & tableName & "'")
found:
    ' Intersect of table range and entirecolumn
    Set getTableColumnRange = intersect(tableR, tableR(1, i).entireColumn)

End Function

Public Function tableLookupCell( _
      tableName As String _
    , columns As collection _
    , values As collection _
    , column As String _
    , Optional WB As Workbook _
    ) As range
    ' Find a cell in a Table and return its range.
    ' The first match is returned.
    '
    ' Args:
    '   TableName: Name of the table.
    '   Columns: Columns to use to search the Values
    '   Values: The values to search for.
    '   Column: Name of any column in the table. Is used to determine the size of the table.
    '   WB: Selected WorkBook
    '
    ' Returns:
    '   The range of the cell that matches its matching Value first.
    
    Set tableLookupCell = intersect(getTableRowRange(tableName, columns, values, WB), getTableColumnRange(tableName, column, WB))

End Function

Private Function miscTableToDicts_EnsureTableDicts(table As Variant, Optional WB As Workbook) As collection
    
    If TypeOf table Is collection Then ' assume if collection, it's already a TableDicts object
        Set miscTableToDicts_EnsureTableDicts = table
    Else
        Set miscTableToDicts_EnsureTableDicts = tableToDicts(CStr(table), WB)
    End If

End Function

Public Function getTableRowIndex( _
      table As Variant _
    , columns As collection _
    , values As collection _
    , Optional WB As Workbook _
    ) As Long
    ' Table can either be a TableToDicts collection, _
      or the name of the table to find
    ' Given a table name, Columns and Values to match _
      this function returns the row in which the first set of values matches
    ' Comparison is case sensitive
    ' If no match is found, SubscriptOutOfRange error is raised
    '
    ' Args:
    '   Table: TableToDicts or name of the table to find.
    '   Columns: Columns to match
    '   Values: Values to match.
    '   WB: Selected WorkBook.
    '
    ' Returns:
    '   The row in which the values matches the comparison.
    
    Dim dict As dictionary
    Dim keyValuePair As collection
    Dim isMatch As Boolean
    Dim rowNumber As Long
    
    For Each dict In miscTableToDicts_EnsureTableDicts(table, WB)
        isMatch = True
        rowNumber = rowNumber + 1
        For Each keyValuePair In zip(columns, values)
            If dict(keyValuePair(1)) <> keyValuePair(2) Then
                isMatch = False
            End If
        Next keyValuePair
        If isMatch = True Then Exit For
    Next dict
    
    If isMatch Then
        getTableRowIndex = rowNumber
    Else
        err.raise errNr.subscriptOutOfRange, , errorMessage(errNr.subscriptOutOfRange, "Columns-values pairs did not find a match")
    End If
    
End Function

Public Sub gotoRowInTable( _
      tableName As String _
    , columns As collection _
    , values As collection _
    , Optional WB As Workbook _
    )
    ' Go to the cell that matches the entry in the Values input.
    '
    ' Args:
    '   TableName: Name of the Table.
    '   Columns: Columns to include in the search.
    '   Values: Values to search for.
    '   WB: Selected WorkBook.
    
    application.GoTo getTableRowRange(tableName, columns, values, WB), True
End Sub

