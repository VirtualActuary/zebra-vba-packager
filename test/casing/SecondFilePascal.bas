Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function SomeFunc(ByVal Col As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   COL: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If Col Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    SomeFunc = Col(1)
    
    For Each Entry In Col
        If Entry < SomeFunc Then
            SomeFunc = Entry
        End If
    Next Entry
End Function
