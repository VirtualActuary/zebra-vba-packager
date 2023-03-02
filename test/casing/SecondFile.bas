Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function SomeFunc(ByVal COL As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   COL: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If COL Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim ENTRY As Variant
    SomeFunc = COL(1)
    
    For Each ENTRY In COL
        If ENTRY < SomeFunc Then
            SomeFunc = ENTRY
        End If
    Next ENTRY
End Function
