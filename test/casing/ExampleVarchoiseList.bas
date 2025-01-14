Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function Min(ByVal cOl As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If cOl Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim enTRY As Variant
    Min = cOl(1)
    
    For Each enTRY In cOl
        If enTRY < Min Then
            Min = enTRY
        End If
    Next enTRY
    
    
    
End Function

Public Function Max(ByVal cOl As Collection) As Variant
    ' Returns the maximum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The maximum value in the collection.
    
    If cOl Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Max = cOl(1)
    Dim enTRY As Variant
    
    For Each enTRY In cOl
        If enTRY > Max Then
            Max = enTRY
        End If
    Next enTRY

End Function
