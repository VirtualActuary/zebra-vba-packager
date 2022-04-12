Attribute VB_Name = "FilePaths"
Option Explicit

' I have never been so confused by a programming language as when I had to write this
' I had so many error and type mismatches w.r.t. array that I really don't know what I did in the
' end or why anything works as it does

Private Sub TestDeleteStringArrayZeroIndex()
    
    Dim arr() As String
    ReDim arr(0 To 2)
    
    arr(0) = "Hello"
    arr(1) = "foo"
    arr(2) = "World"
    
    DeleteArrayElementAt 1, arr
    
    Dim Result As String
    Result = Join(arr, " ")
    
    If Result = "Hello World" Then
        Debug.Print "TestDeleteStringArrayZeroIndex test passed"
    Else
        Debug.Print "TestDeleteStringArrayZeroIndex test failed"
    End If
    
End Sub


Private Sub TestDeleteStringArrayOneIndex()
    
    Dim arr() As String
    ReDim arr(1 To 3)
    
    arr(1) = "Hello"
    arr(2) = "foo"
    arr(3) = "World"
    
    DeleteArrayElementAt 2, arr
    
    Dim Result As String
    Result = Join(arr, " ")
    
    If Result = "Hello World" Then
        Debug.Print "TestDeleteStringArrayOneIndex test passed"
    Else
        Debug.Print "TestDeleteStringArrayOneIndex test failed"
    End If
    
End Sub

Public Sub DeleteArrayElementAt(ByVal index As Integer, ByRef prLst As Variant)
       Dim i As Integer

        ' Move all element back one position
        For i = index + 1 To UBound(prLst)
            prLst(i - 1) = prLst(i)
        Next

        ' Shrink the array by one, removing the last one
        ReDim Preserve prLst(LBound(prLst) To UBound(prLst) - 1)
End Sub


Function absolutePath(ByVal Path As String, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim fso As New FileSystemObject
    
    Path = Replace(Path, "/", "\")
    If Left(Path, 3) = "..\" Or Path = ".." Then
        Path = VLib.GetDirectoryName(wb.Path) & Mid(Path, 3)
    ElseIf Left(Path, 2) = ".\" Or Path = "." Then
        Path = wb.Path & Mid(Path, 2)
    End If
    
    ' Resolve all intermediate ".." and "." paths in order to get the
    ' true Absolute path
    Dim PathVec() As String
    PathVec = VLib.StringSplit(Path, "\")
    
    Dim i As Integer
    i = 0
    While i <= UBound(PathVec)
        If PathVec(i) = "." Then
            DeleteArrayElementAt i, PathVec
            i = i - 1
        ElseIf PathVec(i) = ".." Then
            DeleteArrayElementAt i - 1, PathVec
            DeleteArrayElementAt i - 1, PathVec
            i = i - 2
        End If
        i = i + 1
    Wend
    
    'Why doesn't this work???: VLib.StringJoin(PathVec, "\")
    'Long way around then (sigh)
    Path = ""
    i = 0
    While i <= UBound(PathVec)
       Path = Path & PathVec(i) & "\"
       i = i + 1
    Wend
    If Path <> "" Then Path = Mid(Path, 1, Len(Path) - 1)
    
    absolutePath = Path
    
End Function

