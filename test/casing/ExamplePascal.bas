Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function Min(ByVal Col As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    Min = Col(1)
    
    For Each Entry In Col
        If Entry < Min Then
            Min = Entry
        End If
    Next Entry
    
    
    
End Function

Public Function Max(ByVal Col As Collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Max = Col(1)
    Dim Entry As Variant
    
    For Each Entry In Col
        If Entry > Max Then
            Max = Entry
        End If
    Next Entry

End Function
