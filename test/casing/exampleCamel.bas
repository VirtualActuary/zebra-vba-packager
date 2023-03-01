Attribute VB_Name = "miscCollection"
Option Explicit


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
