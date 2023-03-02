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
    
    For Each Entry In COL
        If EntRY < Min Then
            Min = EnTry
        End If
    Next ENtry
    
    
    
End Function

Public Function Max(ByVal COL As Collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If COL Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Max = COL(1)
    Dim Entry As Variant
    
    For Each Entry In COL
        If Entry > Max Then
            Max = Entry
        End If
    Next Entry

End Function
