Attribute VB_Name = "miscCollection"
Option Explicit


Public Function someFunc(ByVal col As collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   COL: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If col Is Nothing Then
        err.raise number:=91, _
              description:="Collection input can't be empty"
    End If
    
    Dim entry As Variant
    someFunc = col(1)
    
    For Each entry In col
        If entry < someFunc Then
            someFunc = entry
        End If
    Next entry
End Function
