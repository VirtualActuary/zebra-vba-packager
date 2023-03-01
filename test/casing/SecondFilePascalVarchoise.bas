Attribute VB_Name = "MiscCollection"
Option Explicit


Public Function SomeFunc(ByVal coL As Collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   COL: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If coL Is Nothing Then
        Err.Raise Number:=91, _
              Description:="Collection input can't be empty"
    End If
    
    Dim Entry As Variant
    SomeFunc = coL(1)
    
    For Each Entry In coL
        If Entry < SomeFunc Then
            SomeFunc = Entry
        End If
    Next Entry
End Function
