Attribute VB_Name = "MiscCollection"
Option Explicit

#If VBA7 Then
    #If Win64 Then
        Declare PtrSafe Function GetKeyState Lib "User32" (ByVal VKey As Integer) As Integer
    #Else
        Declare Function GetKeyState Lib "User32" (ByVal VKey As Integer) As Integer
    #End If
#Else
    Declare Function GetKeyState Lib "User32" (ByVal VKey As Integer) As Integer
#End If


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
